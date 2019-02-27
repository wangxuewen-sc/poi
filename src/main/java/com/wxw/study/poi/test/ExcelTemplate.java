package com.wxw.study.poi.test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections.CollectionUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.CellRangeAddressList;

/**
 * Created by hqk2015@foxmail.com on 2017/11/14.
 * poi version: poi-3.10-FINAL-20140208
 */
public class ExcelTemplate {
    private static final Logger LOG = Logger.getLogger(ExcelTemplate.class);
    private static final int XLS_MAX_ROW = 65535; //0��ʼ
    private static final String MAIN_SHEET_NAME = "template";
    private static final String HIDDEN_SHEET1_NAME = "hidden1";
    private static final String HIDDEN_SHEET2_NAME = "hidden2";
    private static final String WAREHOUSE_NAMES = "warehouses";
    private static final String DEVICE_NAMES = "devices";
    private static final String DEVICE_TYPE_NAMES = "deviceTypes";


    /**
     * �������excelģ��
     * @param filePath
     * @param headers
     * @param devices
     * @param deviceTypes
     * @param warehouses
     * @param warehouseAndShelves
     * @return
     * @throws IOException
     */
    private static File createStoreInExcelTemplate(String filePath, List<String> headers,List<String> devices, List<String> deviceTypes, List<String> warehouses, Map<String, List<String>> warehouseAndShelves) throws IOException {
        FileOutputStream out = null;
        File file;
        try {
            file = new File(filePath); //д�ļ�
//            file = File.createTempFile("�������ģ��", ".xls");
            out = new FileOutputStream(file);
            HSSFWorkbook wb = new HSSFWorkbook();//����������
            HSSFSheet mainSheet = wb.createSheet(MAIN_SHEET_NAME);
            HSSFSheet dtHiddenSheet = wb.createSheet(HIDDEN_SHEET1_NAME);
            HSSFSheet wsHiddenSheet = wb.createSheet(HIDDEN_SHEET2_NAME);
//            wb.setSheetHidden(1, true); //���ڶ������ڴ洢���������ݵ�sheet����
//            wb.setSheetHidden(2, true);
            initHeaders(wb, mainSheet, headers);
            initDevicesAndType(wb, dtHiddenSheet, devices, deviceTypes);
            initWarehousesAndShelves(wb, wsHiddenSheet, warehouses, warehouseAndShelves);
            initSheetNameMapping(mainSheet);
            initOtherConstraints(mainSheet);
            out.flush();
            wb.write(out);
//            file.deleteOnExit();
        } catch (Exception e) {
            e.printStackTrace();
            LOG.error("�������excelģ��ʧ�ܣ�", e);
            throw e;
        } finally {
            if (out != null)
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
        }
        return file;
    }

    /**
     * ��ʼ���ֿ�&��������������
     *
     * @param workbook
     * @param wsSheet
     * @param warehouses
     * @param warehousesAndShelves
     */
    private static void initWarehousesAndShelves(HSSFWorkbook workbook, HSSFSheet wsSheet, List<String> warehouses, Map<String, List<String>> warehousesAndShelves) {
        writeWarehouses(workbook, wsSheet, warehouses);
        writeShelves(workbook, wsSheet, warehouses, warehousesAndShelves);
        initWarehouseNameMapping(workbook, wsSheet.getSheetName(), warehouses.size());
    }

    /**
     * ������sheet��д���������������
     *
     * @param workbook
     * @param wsSheet
     * @param warehouses
     * @param warehousesAndShelves
     */
    private static void writeShelves(HSSFWorkbook workbook, HSSFSheet wsSheet, List<String> warehouses, Map<String, List<String>> warehousesAndShelves) {
        for (int i = 0; i < warehouses.size(); i++) {
            int referColNum = i + 1;
            String warehouseName = warehouses.get(i);
            List<String> shelves = warehousesAndShelves.get(warehouseName);
            if (CollectionUtils.isNotEmpty(shelves)) {
                int rowCount = wsSheet.getLastRowNum();
                if(rowCount == 0 && wsSheet.getRow(0) == null )
                    wsSheet.createRow(0);
                for (int j = 0; j < shelves.size(); j++) {
                    if (j <= rowCount) { //ǰ�洴�������У�ֱ�ӻ�ȡ�У�������
                        wsSheet.getRow(j).createCell(referColNum).setCellValue(shelves.get(j)); //���ö�Ӧ��Ԫ���ֵ
                    } else { //δ���������У�ֱ�Ӵ����С�������
                        wsSheet.setColumnWidth(j, 4000); //����ÿ�е��п�
                        //�����С�������
                        wsSheet.createRow(j).createCell(referColNum).setCellValue(shelves.get(j)); //���ö�Ӧ��Ԫ���ֵ
                    }
                }
            }
            initShelfNameMapping(workbook, wsSheet.getSheetName(), warehouseName, referColNum, shelves.size());
        }
    }

    /**
     * ������ⷿ����ѡ�����ݹ���
     *
     * @param workbook
     * @param wsSheetName
     * @param warehouseName
     * @param referColNum
     * @param shelfQuantity
     */
    private static void initShelfNameMapping(HSSFWorkbook workbook, String wsSheetName, String warehouseName, int referColNum, int shelfQuantity) {
        Name name = workbook.createName();
        // ���û�������
        name.setNameName(warehouseName);
        String referColName = getColumnName(referColNum);
        String formula = wsSheetName + "!$" + referColName + "$1:$" + referColName + "$" + shelfQuantity;
        name.setRefersToFormula(formula);
    }

    /**
     * ������sheet��д��ⷿ����������
     *
     * @param workbook
     * @param wsSheet
     * @param warehouses
     */
    private static void writeWarehouses(HSSFWorkbook workbook, HSSFSheet wsSheet, List<String> warehouses) {
        for (int i = 0; i < warehouses.size(); i++) {
            HSSFRow row = wsSheet.createRow(i);
            HSSFCell cell = row.createCell(0);
            cell.setCellValue(warehouses.get(i));
        }
    }

    /**
     * ��ʼ���ⷿѡ�����ơ�
     *
     * @param workbook
     * @param wsSheetName
     * @param warehouseQuantity
     */
    private static void initWarehouseNameMapping(HSSFWorkbook workbook, String wsSheetName, int warehouseQuantity) {
        Name name = workbook.createName();
        // ���òֿ�����
        name.setNameName(WAREHOUSE_NAMES);
        name.setRefersToFormula(wsSheetName + "!$A$1:$A$" + warehouseQuantity);
    }

    /**
     * ��������ֵȷ����Ԫ��λ�ã����磺0-A, 27-AB��
     *
     * @param index
     * @return
     */
    public static String getColumnName(int index) {
        StringBuilder s = new StringBuilder();
        while (index >= 26) {
            s.insert(0, (char) ('A' + index % 26));
            index = index / 26 - 1;
        }
        s.insert(0, (char) ('A' + index));
        return s.toString();
    }

    private static void initDevicesAndType(HSSFWorkbook wb, HSSFSheet dtHiddenSheet, List<String> devices, List<String> deviceTypes) {
        writeDevices(wb, dtHiddenSheet, devices);
        writeDeviceTypes(wb, dtHiddenSheet, deviceTypes);
        initDevicesNameMapping(wb, dtHiddenSheet.getSheetName(), devices.size());
        initDeviceTypesNameMapping(wb, dtHiddenSheet.getSheetName(), deviceTypes.size());
    }

    private static void initDeviceTypesNameMapping(HSSFWorkbook wb, String dtHiddenSheetName, int deviceTypeQuantity) {
        Name name = wb.createName();
        // �����豸�����ơ�
        name.setNameName(DEVICE_TYPE_NAMES);
        name.setRefersToFormula(dtHiddenSheetName + "!$B$1:$B$" + deviceTypeQuantity);
    }

    private static void writeDeviceTypes(HSSFWorkbook wb, HSSFSheet dtHiddenSheet, List<String> deviceTypes) {
        int lastRow = dtHiddenSheet.getLastRowNum();
        if(lastRow == 0 && dtHiddenSheet.getRow(0) == null )
            dtHiddenSheet.createRow(0);
        if (CollectionUtils.isNotEmpty(deviceTypes))
            for (int j = 0; j < deviceTypes.size(); j++) {
                if (j <= lastRow) { //ǰ�洴�������У�ֱ�ӻ�ȡ�У�������
                    dtHiddenSheet.getRow(j).createCell(1).setCellValue(deviceTypes.get(j)); //���ö�Ӧ��Ԫ���ֵ
                } else { //δ���������У�ֱ�Ӵ����С�������
                    dtHiddenSheet.setColumnWidth(j, 4000); //����ÿ�е��п�
                    //�����С�������
                    dtHiddenSheet.createRow(j).createCell(1).setCellValue(deviceTypes.get(j)); //���ö�Ӧ��Ԫ���ֵ
                }
            }
    }

    private static void writeDevices(HSSFWorkbook wb, HSSFSheet dtHiddenSheet, List<String> devices) {
        for (int i = 0; i < devices.size(); i++) {
            HSSFRow row = dtHiddenSheet.createRow(i);
            HSSFCell cell1 = row.createCell(0);
            cell1.setCellValue(devices.get(i));
        }
    }

    private static void initDevicesNameMapping(HSSFWorkbook workbook, String dtHiddenSheetName, int deviceQuantity) {
        Name name = workbook.createName();
        // �����豸�����ơ�
        name.setNameName(DEVICE_NAMES);
        name.setRefersToFormula(dtHiddenSheetName + "!$A$1:$A$" + deviceQuantity);
    }

    /**
     * ��ʼ����ͷ
     *
     * @param wb
     * @param mainSheet
     * @param headers
     */
    private static void initHeaders(HSSFWorkbook wb, HSSFSheet mainSheet, List<String> headers) {
        //��ͷ��ʽ
        HSSFCellStyle style = wb.createCellStyle();
//        style.setAlignment(style.getAlignmentEnum().CENTER); // ����һ�����и�ʽ
        //������ʽ
        HSSFFont fontStyle = wb.createFont();
        fontStyle.setFontName("΢���ź�");
        fontStyle.setFontHeightInPoints((short) 12);
//        fontStyle.setBoldweight();
        style.setFont(fontStyle);
        //����sheet1����
        HSSFRow rowFirst = mainSheet.createRow(0);//��һ��sheet�ĵ�һ��Ϊ����
        mainSheet.createFreezePane(0, 1, 0, 1); //�����һ��
        //д����
        for (int i = 0; i < headers.size(); i++) {
            HSSFCell cell = rowFirst.createCell(i); //��ȡ��һ�е�ÿ����Ԫ��
            mainSheet.setColumnWidth(i, 4000); //����ÿ�е��п�
            cell.setCellStyle(style); //����ʽ
            cell.setCellValue(headers.get(i)); //����Ԫ����д����
        }
    }

    /**
     * ��sheet���������ʼ��
     *
     * @param mainSheet
     */
    private static void initSheetNameMapping(HSSFSheet mainSheet) {
        DataValidation warehouseValidation = getDataValidationByFormula(WAREHOUSE_NAMES, 3);
        DataValidation shelfValidation = getDataValidationByFormula("INDIRECT($D1)", 4); //formulaͬ"INDIRECT(INDIRECT(\"R\"&ROW()&\"C\"&(COLUMN()-1),FALSE))"
        DataValidation deviceValidation = getDataValidationByFormula(DEVICE_NAMES, 0);
        DataValidation deviceTypeValidation = getDataValidationByFormula(DEVICE_TYPE_NAMES, 1);
        // ��sheet�����֤����
        mainSheet.addValidationData(warehouseValidation);
        mainSheet.addValidationData(shelfValidation);
        mainSheet.addValidationData(deviceValidation);
        mainSheet.addValidationData(deviceTypeValidation);
    }

    /**
     * ��ʼ�������������޶�
     *
     * @param mainSheet
     */
    private static void initOtherConstraints(HSSFSheet mainSheet) {
        DataValidation quantityValidation = getDecimalValidation(1, XLS_MAX_ROW, 2);
        DataValidation tierValidation = getDecimalValidation(1, XLS_MAX_ROW, 5);
        DataValidation colValidation = getDecimalValidation(1, XLS_MAX_ROW, 6);
        mainSheet.addValidationData(tierValidation);
        mainSheet.addValidationData(colValidation);
        mainSheet.addValidationData(quantityValidation);
    }
    /**
     * ������������ʾ
     *
     * @param formulaString
     * @param columnIndex
     * @return
     */
    public static DataValidation getDataValidationByFormula(String formulaString, int columnIndex) {
        // ���������б�����
        DVConstraint constraint = DVConstraint.createFormulaListConstraint(formulaString);
        // ����������Ч�Լ������ĸ���Ԫ���ϡ�
        // �ĸ������ֱ��ǣ���ʼ�С���ֹ�С���ʼ�С���ֹ��
        CellRangeAddressList regions = new CellRangeAddressList(1, XLS_MAX_ROW, columnIndex, columnIndex);
        // ������Ч�Զ���
        DataValidation dataValidationList = new HSSFDataValidation(regions, constraint);
        dataValidationList.createErrorBox("Error", "��ѡ���������Ч��ѡ�����������ģ�����ԣ�");
        String promptText = initPromptText(columnIndex);
        dataValidationList.createPromptBox("", promptText);
        return dataValidationList;
    }

    /**
     * ����������֤
     * @param firstRow
     * @param lastRow
     * @param columnIndex
     * @return
     */
    private static DataValidation getDecimalValidation(int firstRow, int lastRow, int columnIndex) {
        // ����һ������>0������
        DVConstraint constraint = DVConstraint.createNumericConstraint(DVConstraint.ValidationType.INTEGER, DVConstraint.OperatorType.GREATER_OR_EQUAL, "0", "0");
        // �趨���ĸ���Ԫ����Ч
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, columnIndex, columnIndex);
        // �����������
        HSSFDataValidation decimalVal = new HSSFDataValidation(regions, constraint);
        decimalVal.createPromptBox("", initPromptText(columnIndex));
        decimalVal.createErrorBox("����ֵ���ͻ��С����", "��ֵ�ͣ����������0 ��������");
        return decimalVal;
    }

    /**
     * ��ʼ����������ʾ��Ϣ
     *
     * @param columnIndex
     * @return
     */
    private static String initPromptText(int columnIndex) {
        String promptText ="";
        //custom column prompt
        switch (columnIndex) {
            case 2:
                promptText = "���������0��������";
                break;
            case 4:
                promptText = "������ѡ���������Ч�����ѡ��ⷿ��";
                break;
            case 5:
                promptText = "���������0��������";
                break;
            case 6:
                promptText = "���������0��������";
                break;
        }
        return promptText;
    }

    public static File test() throws IOException {
        //************* mock data ************
        List<String> headers = Arrays.asList("�豸����", "������", "����", "��ſⷿ", "��Ż���", "��Ų�", "�����");
        List<String> deviceTypes = Arrays.asList("type1", "type2", "type3", "type4");
        List<String> devices = Arrays.asList("�豸1", "�豸2", "�豸3");
        List<String> warehouses = Arrays.asList("�ⷿ1", "�ⷿ2");
        Map<String, List<String>> warehousesAndShelves = new HashMap<>();
        warehousesAndShelves.put("�ⷿ1", Arrays.asList("����1-1", "����1-2", "����1-3"));
        warehousesAndShelves.put("�ⷿ2", Arrays.asList("����2-1", "����2-2", "����2-3"));
        return ExcelTemplate.createStoreInExcelTemplate("E:/test.xls", headers,devices, deviceTypes, warehouses, warehousesAndShelves);
    }

    public static void main(String[] args) throws IOException {
        ExcelTemplate.test();
    }
}
