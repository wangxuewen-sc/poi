package com.wxw.util.excelIo;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 
 * @author WangXuewen
 *
 * @param <T>
 */
public class Excel<T> {
	private Class<T> clasz;
	private int excelAnnoCount;
	private String sheetName;
//	����EXCEL�ļ�
	XSSFWorkbook workBook;
	
	/**
	 * @param clasz װ���ݵĶ�����
	 */
	public Excel(Class<T> clasz) {
		this.clasz = clasz;
	}
	
	/**
	 * @param clasz װ���ݵĶ�����
	 * @param sheetName ����
	 */
	public Excel(Class<T> clasz,String sheetName) {
		this.clasz = clasz;
		this.sheetName = sheetName;
	}
	
	/**
	 * @return ����һ����������ע��ֵ��Map����
	 */
	private Map<String,EntityFieldMap> getEntityFieldMap(){
		Map<String,EntityFieldMap> entityMap = new HashMap<>();
		Field[] fields = this.clasz.getDeclaredFields();
		int zeroCount = 0;
		for(Field field:fields) {
			if(field.isAnnotationPresent(ExcelProperty.class)) {
				ExcelProperty excelProp = this.getAnnoVal(field);
				if(excelProp.sequence()==0)zeroCount++;
				entityMap.put(field.getName(),
					new EntityFieldMap(field.getName(),
					excelProp.zh(),
					excelProp.sequence()==0 && zeroCount>1?this.excelAnnoCount:excelProp.sequence(),
					excelProp.requisite()));
				this.excelAnnoCount++;
			}
		}
		return entityMap;
	}
	
	/**
	 * xlsx
	 * Excel�����ͷ��˳������ԵĶ�Ӧ
	 * @param row
	 */
	private Map<Integer,EntityFieldMap> getExcelTitleSort(XSSFRow row){
		int lastCell = row.getLastCellNum();
		Map<String,EntityFieldMap> entityFieldMap = this.getEntityFieldMap();
		Map<Integer,EntityFieldMap> importEntityFieldMap = new HashMap<>();
		for(int i=0;i<lastCell;i++) {
			XSSFCell cell = row.getCell(i);
			for(Entry<String,EntityFieldMap> entry:entityFieldMap.entrySet()) {
				if(entry.getValue().getZh().equals(cell.getStringCellValue())) {
					importEntityFieldMap.put(i, entry.getValue());
					continue;
				}
			}
		}
		return importEntityFieldMap;
	}
	
	/**
	 * xls
	 * Excel�����ͷ��˳������ԵĶ�Ӧ
	 * @param row
	 */
	private Map<Integer,EntityFieldMap> getExcelTitleSort(HSSFRow row){
		int lastCell = row.getLastCellNum();
		Map<String,EntityFieldMap> entityFieldMap = this.getEntityFieldMap();
		Map<Integer,EntityFieldMap> importEntityFieldMap = new HashMap<>();
		for(int i=0;i<lastCell;i++) {
			HSSFCell cell = row.getCell(i);
			importEntityFieldMap.put(i, entityFieldMap.get(cell.getStringCellValue()));
		}
		return importEntityFieldMap;
	}
	
	/**
	 * 
	 * @param field ���Զ���
	 * @return ����һ�� ExcelProperty ����
	 * �÷�����ȡ����ָ�������ϵ�ָ��ע�����
	 */
	private ExcelProperty getAnnoVal(Field field) {
		field.setAccessible(true);
		ExcelProperty excelProp = field.getAnnotation(ExcelProperty.class);
		field.setAccessible(false);
		return excelProp;
	}
	
	/***
	 * 
	 * @param ins ����һ��InputStream ����
	 * @param file ����һ������Ҫ��ȡ��EXCEL ��File ����
	 * @return
	 * @throws IOException
	 */
	public List<T> readExcel(InputStream ins){
		List<T> datas = new ArrayList<T>();
			ins = FileMagic.prepareToCheckMagic(ins);
			FileMagic fm = null;
			try {
				fm = FileMagic.valueOf(ins);
			} catch (IOException e1) {
				e1.printStackTrace();
			}
			switch (fm) {
	        case OLE2:
	            datas =this.readExcel2003(ins);break;
	        case OOXML:
	            datas = this.readExcel2007(ins);break;
	        default:
	            try {
					throw new InvalidFormatException("Your InputStream was neither an OLE2 stream, nor an OOXML stream");
				} catch (InvalidFormatException e) {
					e.printStackTrace();
				}
	        }
		return datas;
	}
	
	private List<T> readExcel2003(InputStream ins) {
//		ʹ�ù���������洢�ļ���
		HSSFWorkbook workBook = null;
//			������Ԫ��
		List<T> datas =new ArrayList<T>();
		try {
			workBook = new HSSFWorkbook(ins);
//			��ȡָ��������
			HSSFSheet sheet = this.sheetName==null?workBook.getSheetAt(0):workBook.getSheet(this.sheetName);
			if(sheet==null) {
				throw new FileNotFoundException("δ�ҵ���Ϊ��"+this.sheetName+"����sheet��!");
			}
			//��ȡ��ͷ��˳��ӳ��
			Map<Integer,EntityFieldMap> entityMap = getExcelTitleSort(sheet.getRow(0));
//			��ȡ���һ���к�
			int lastRowNum = sheet.getLastRowNum();
			for(int i=1;i<=lastRowNum;i++) {
				HSSFRow row = sheet.getRow(i);
				if(row==null)break;
				datas.add(i,this.getRowData(row, entityMap));
			}
		} catch (IOException e) {
			e.printStackTrace();
		}finally {
			try {
				if(workBook!=null)workBook.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

		return datas;
	}


	private List<T> readExcel2007(InputStream ins){
//		ʹ�ù���������洢�ļ���
		XSSFWorkbook workBook = null;
//		������Ԫ��
		List<T> datas = null;
		try {
			workBook = new XSSFWorkbook(ins);
//			��ȡָ��������
			XSSFSheet sheet = this.sheetName==null?workBook.getSheetAt(0):workBook.getSheet(sheetName);
//			��ȡ���һ���к�
			if(sheet==null) {
				throw new FileNotFoundException("δ�ҵ���Ϊ��"+this.sheetName+"����sheet��!");
			}
			int lastRowNum = sheet.getLastRowNum();
			//��ȡ��ͷ��˳��ӳ��
			Map<Integer,EntityFieldMap> entityMap = getExcelTitleSort(sheet.getRow(0));
			datas=new ArrayList<T>(entityMap.size());
			//ȥ����ͷ��0�У��ӵ�1�п�ʼ
			for(int i=1;i<=lastRowNum;i++) {
				XSSFRow row = sheet.getRow(i);
				if(row==null)break;
				try {
					datas.add(getRowData(row,entityMap));
				} catch (InstantiationException |
						IllegalAccessException | 
						NoSuchFieldException | 
						SecurityException e) {
					e.printStackTrace();
				}
			}
		} catch (IOException e) {
			e.printStackTrace();
		}finally {
			
				try {
					if(ins!=null)ins.close();
					if(workBook!=null)workBook.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			
		}
		return datas;
	}
	
	private T getRowData(XSSFRow row,Map<Integer,EntityFieldMap> entityMap) throws InstantiationException, IllegalAccessException, NoSuchFieldException, SecurityException {
		int lastCellNum = row.getLastCellNum();//���һ�б��
		T t = clasz.newInstance();
		for(int j=0;j<=lastCellNum;j++) {
			XSSFCell cell = row.getCell(j);
			if(cell==null)break;
			
			Field field = clasz.getDeclaredField(entityMap.get(j).getFiledName());
			field.setAccessible(true);
			this.setFieldValue(t, field, cell);
			field.setAccessible(false);
		}
		return t;
	}
	
	private T getRowData(HSSFRow row,Map<Integer,EntityFieldMap> entityMap) {
		int lastCellNum = row.getLastCellNum();//���һ�б��
		
		T t = null;
		try {
			t = clasz.newInstance();
			for(int j=0;j<=lastCellNum;j++) {
				HSSFCell cell = row.getCell(j);
				if(cell==null)break;
				
				Field field = clasz.getField(entityMap.get(j).getFiledName());
				field.setAccessible(true);
				this.setFieldValue(t, field, cell);
				field.setAccessible(false);
			}
		} catch (InstantiationException e) {
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		} catch (NoSuchFieldException e) {
			e.printStackTrace();
		} catch (SecurityException e) {
			e.printStackTrace();
		}finally {
			
		}
		
		return t;
	}
	
	//��������ֵ
	private void setFieldValue(T t,Field field,XSSFCell cell) throws IllegalArgumentException, IllegalAccessException {

		Class<?> clasz = field.getType();
		String value = this.getCellType(cell);
		field.setAccessible(true);
		
		if(clasz.equals(String.class)) {
			field.set(t, value);
		}else if(clasz.equals(double.class) || clasz.equals(Double.class)) {
			field.set(t,Double.valueOf(value));
		}else if(clasz.equals(int.class) || clasz.equals(Integer.class)) {
			field.set(t,Integer.valueOf(value));
		}else if(clasz.equals(boolean.class) || clasz.equals(Boolean.class)) {
			field.set(t,Boolean.valueOf(value));
		}else if(clasz.equals(long.class) || clasz.equals(Long.class)){
			field.set(t, Long.valueOf(value));
		}else if(clasz.equals(short.class) || clasz.equals(Short.class)) {
			field.setShort(t, Short.valueOf(value));
		}else if(clasz.equals(byte.class) || clasz.equals(Byte.class)) {
			field.setByte(t, Byte.valueOf(value));
		}else if(clasz.equals(float.class) || clasz.equals(Float.class)) {
			field.setFloat(t, Float.valueOf(value));
		}else if(clasz.equals(Date.class)) {
			try {
				field.set(t, new SimpleDateFormat("yyyy/MM/dd hh:mm:ss").parse(value));
			} catch (ParseException e) {
				e.printStackTrace();
			}
		}else if(clasz.equals(char.class)||Character.class.equals(clasz)) {
			if(value.length()>0 && value!="")
			field.setChar(t, Character.valueOf(value.charAt(0)));
		}
		
		field.setAccessible(false);
	}
	
	private void setFieldValue(T t,Field field,HSSFCell cell) throws IllegalArgumentException, IllegalAccessException {

		Class<?> clasz = field.getType();
		String value = this.getCellType(cell);
		field.setAccessible(true);
		
		if(clasz.equals(String.class)) {
			field.set(t, value);
		}else if(clasz.equals(double.class) || clasz.equals(Double.class)) {
			field.set(t,Double.valueOf(value));
		}else if(clasz.equals(int.class) || clasz.equals(Integer.class)) {
			field.set(t,Integer.valueOf(value));
		}else if(clasz.equals(boolean.class) || clasz.equals(Boolean.class)) {
			field.set(t,Boolean.valueOf(value));
		}else if(clasz.equals(long.class) || clasz.equals(Long.class)){
			field.set(t, Long.valueOf(value));
		}else if(clasz.equals(short.class) || clasz.equals(Short.class)) {
			field.setShort(t, Short.valueOf(value));
		}else if(clasz.equals(byte.class) || clasz.equals(Byte.class)) {
			field.setByte(t, Byte.valueOf(value));
		}else if(clasz.equals(float.class) || clasz.equals(Float.class)) {
			field.setFloat(t, Float.valueOf(value));
		}else if(clasz.equals(Date.class)) {
			try {
				field.set(t, new SimpleDateFormat("yyyy/MM/dd hh:mm:ss").parse(value));
			} catch (ParseException e) {
				e.printStackTrace();
			}
		}else if(clasz.equals(char.class)||Character.class.equals(clasz)) {
			if(value.length()>0 && value!="")
			field.setChar(t, Character.valueOf(value.charAt(0)));
		}
		
		field.setAccessible(false);
	}
	
	//��ȡ��Ԫ����������
	private String getCellType(XSSFCell cell) {
		SimpleDateFormat dateFormat = new SimpleDateFormat();
		String strVal = null;
		String styleStr =cell.getCellStyle().getDataFormatString();
		switch(cell.getCellTypeEnum()) {
		case NUMERIC:
			if(org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
				dateFormat.applyPattern("yyyy/MM/dd hh:mm:ss");
				strVal = dateFormat.format(cell.getDateCellValue());
			}else if("0.00E+00".equals(styleStr)) {
				strVal = new BigDecimal(cell.getNumericCellValue()).toPlainString();
			}else if("General".equals(styleStr)){
				strVal = new BigDecimal(cell.getNumericCellValue()).toPlainString();
			}else {
				strVal = String.valueOf(cell.getNumericCellValue());
			}
			;break;
		case STRING:
			strVal = cell.getStringCellValue();
			break;
		case BOOLEAN:
			strVal = String.valueOf(cell.getBooleanCellValue());
			;break;
		case FORMULA:
			strVal = String.valueOf(cell.getNumericCellValue());
			;break;
		case BLANK:
			strVal = String.valueOf(cell.getNumericCellValue());;
			;break;
		case ERROR:
			throw new RuntimeException("��ȡ�������ʹ��󣡣���");
		default:
			throw new RuntimeException("δ֪�������ͣ�����");
		}
		return strVal;
	}
	
	//��ȡ��Ԫ����������
	private String getCellType(HSSFCell cell) {
		SimpleDateFormat dateFormat = new SimpleDateFormat();
		String strVal = null;
		String styleStr =cell.getCellStyle().getDataFormatString();
		switch(cell.getCellTypeEnum()) {
		case NUMERIC:
			if(HSSFDateUtil.isCellDateFormatted(cell)) {
				dateFormat.applyPattern("yyyy/MM/dd hh:mm:ss");
				strVal = dateFormat.format(cell.getDateCellValue());
			}else if("0.00E+00".equals(styleStr)) {
				strVal = new BigDecimal(cell.getNumericCellValue()).toPlainString();
			}else if("General".equals(styleStr)){
				strVal = new BigDecimal(cell.getNumericCellValue()).toPlainString();
			}else {
				strVal = String.valueOf(cell.getNumericCellValue());
			}
			;break;
		case STRING:
			strVal = cell.getStringCellValue();
			break;
		case BOOLEAN:
			strVal = String.valueOf(cell.getBooleanCellValue());
			;break;
		case FORMULA:
			strVal = String.valueOf(cell.getNumericCellValue());
			;break;
		case BLANK:
			strVal = String.valueOf(cell.getNumericCellValue());;
			;break;
		case ERROR:
			throw new RuntimeException("��ȡ�������ʹ��󣡣���");
		default:
			throw new RuntimeException("δ֪�������ͣ�����");
		}
		return strVal;
	}
	
	
	/**
	 * ���ݴ���Ķ��󣬴���Excel�ļ���ת��Ϊ�����
	 * @param objList ����һ�����ݶ��������excel������
	 */
	public XSSFWorkbook createExcel(List<T> objList){
		if(this.sheetName==null)this.sheetName="sheet";
		this.workBook = new XSSFWorkbook();
		XSSFCellStyle contextStyle = this.workBook.createCellStyle();
//		����һ�����������ñ���
		XSSFSheet sheet = this.workBook.createSheet(this.sheetName);
//		�ڹ������ϴ���һ����
		XSSFRow row = sheet.createRow(0);
//		�����ϴ�����Ԫ��
		Map<String, EntityFieldMap> map =this.getEntityFieldMap();
//		������ͷ
		this.createTableTilte(row,map);
//		����������
		this.creatDataRows(sheet,objList,map,contextStyle);
//		����Ĭ�ϻ״̬��
		this.workBook.setActiveSheet(0);
//		����ļ�
		return this.workBook;
		
	}
	

	/**
	 * ������ͷ
	 * @param row
	 * @param map
	 */
	private void createTableTilte(XSSFRow row,Map<String, EntityFieldMap> map) {
		for(Entry<String, EntityFieldMap> entry:map.entrySet()) {
			row.createCell(entry.getValue().getSequence()).setCellValue(entry.getValue().getZh());
		}
	}
	
	/**
	 * ����������
	 * @param sheet
	 * @param objList
	 */
	private void creatDataRows(XSSFSheet sheet, List<T> objList,Map<String, EntityFieldMap> map,XSSFCellStyle cellStyle) {
		for(int rowIndex = 1;rowIndex<=objList.size();rowIndex++) {
			XSSFRow row = sheet.createRow(rowIndex);//����������
			T t = objList.get(rowIndex-1);
			Field field = null;
			for(Entry<String, EntityFieldMap> entry:map.entrySet()) {
				try {
					field = this.clasz.getDeclaredField(entry.getKey());//������������ȡ����
					field.setAccessible(true);
					XSSFCell cell= row.createCell(entry.getValue().getSequence());
					//����ע�����õ�Ԫ���ʽ
					DataType dt = field.getAnnotation(ExcelProperty.class).dataType();
					if(!DataType.NULL.equals(dt)) {
						if(DataType.STRING.equals(dt)) {
							cell.setCellType(CellType.STRING);
						}else if(DataType.NUMBER.equals(dt)){
							cell.setCellType(CellType.NUMERIC);
						}
					}
					//���������б�
					if(field.isAnnotationPresent(Select.class)) {
						this.createSelectData(field,sheet,rowIndex,entry.getValue().getSequence());
					}else {
						cell.setCellValue(formatObjValueToStr(field,field.get(t)));
					}
						
					
				} catch (NoSuchFieldException | SecurityException | IllegalArgumentException | IllegalAccessException e) {
					e.printStackTrace();
				}finally {
					field.setAccessible(false);
				}
			}
		}
	}
	
//	/**
//	 * 
//	 */
//	public DataValidation getDataValidationByFormula(String formulaString,int firstRow,int lastRow,int firstCol, int lastCol) {
//        // ���������б�����
//        DVConstraint constraint = DVConstraint.createFormulaListConstraint(formulaString);
//        // ����������Ч�Լ������ĸ���Ԫ���ϡ�
//        // �ĸ������ֱ��ǣ���ʼ�С���ֹ�С���ʼ�С���ֹ��
//        CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
//        // ������Ч�Զ���
//        DataValidation dataValidationList = new HSSFDataValidation(regions, constraint);
//        dataValidationList.createErrorBox("Error", "��ѡ���������Ч��ѡ�����������ģ�����ԣ�");
//        return dataValidationList;
//    }
	
	/**
	 * ���������б�˵�
	 */
	private void createSelectData(Field field,XSSFSheet sheet,int rowIndex,int cellIndex) {
		Select select =field.getAnnotation(Select.class);
		Class<?> clazz = select.clazz();
		Method method = null;
		Object obj = null;
		List<?> values = null;
		Map<?,?> cascadeValue = null;
		try {
			System.out.println("���÷�����"+select.method());
			method= clazz.getMethod(select.method());
			obj = clazz.newInstance();
			if(select.cascadeNumber()==0) {
				values = (List<?>) method.invoke(obj, new Object[] {});
				this.setSelectMenum(values, sheet, rowIndex, cellIndex);
			}else {
				cascadeValue = (Map<?, ?>) method.invoke(obj, new Object[] {});
				String hiddenSheetName = "sheet_"+field.getName();
				XSSFSheet hiddenSheet = this.workBook.createSheet(hiddenSheetName);
				this.createCascade(field, cascadeValue, sheet, hiddenSheet, rowIndex, cellIndex);
			}
		} catch (NoSuchMethodException e) {
			e.printStackTrace();
		} catch (InstantiationException e) {
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		} catch (IllegalArgumentException e) {
			e.printStackTrace();
		} catch (InvocationTargetException e) {
			e.printStackTrace();
		}
	}
	
	
	public void createCascade(Field field,Map<?,?> cascadeValue,XSSFSheet sheet,XSSFSheet hiddenSheet,int rowIndex,int cellIndex) {
		int rownum = 0;
		int cellnum = 0;
		int rows;
		XSSFRow row;
		for(Entry<?, ?> entry:cascadeValue.entrySet()) {
			rows = hiddenSheet.getLastRowNum();
			System.out.println(rows);
			if(rows==0 && rows<=rownum) {
				row = hiddenSheet.createRow(rownum);
			}else {
				row = hiddenSheet.getRow(rownum);
			}
			
			XSSFCell cell = row.createCell(0);
			cell.setCellValue((String)entry.getKey());
			List<?> list = (List<?>) entry.getValue();
			for(int i=0;i<list.size();i++) {
				if(list.get(i).getClass().equals(Map.class)) {
					
				}else {
					if(hiddenSheet.getLastRowNum()>=i) {
						row = hiddenSheet.getRow(i);
					}else {
						row = hiddenSheet.createRow(i);
					}
					XSSFCell cell1 = row.createCell(cellnum+1);
					cell1.setCellValue((String)list.get(i));
				}
			}
			rownum++;
			cellnum++;
		}
//		��Ҫ�õ�������map�Ĵ�С��ÿһ��list�Ĵ�С��A=65,
//		�ٴ���һ������ҳ
//		�ڰ�map��keyֵ�ŵ���һ��
//		DataValidation Lv1 = this.getDataValidationByFormula(field.getName(), 0, cascadeValue.size(), 1, 1);
		
	}
	
	/**
	 * ��������ѡ������
	 * @param values
	 * @param sheet
	 * @param rowIndex
	 * @param cellIndex
	 */
	public void setSelectMenum(List<?> values,XSSFSheet sheet,int rowIndex,int cellIndex) {
		// ���������б�����
//      DVConstraint constraint = DVConstraint.createFormulaListConstraint(formulaString);
		// ����������Ч�Լ������ĸ���Ԫ���ϡ�
      // �ĸ������ֱ��ǣ���ʼ�С���ֹ�С���ʼ�С���ֹ��
      String[] str = new String[values.size()];
      XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
      XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
	            .createExplicitListConstraint(values.toArray(str));
	    
	    CellRangeAddressList addressList = new CellRangeAddressList(rowIndex, rowIndex, cellIndex, cellIndex);
	    XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(
	                dvConstraint, addressList);
	    validation.createErrorBox("����", "��ѡ�������б��е�ֵ");
	    validation.setSuppressDropDownArrow(true);
	    validation.setShowErrorBox(true);//��ʾ�Ƿ�������ʾ
      sheet.addValidationData(validation);
	}
	
	/****
	 * ���������ݶ�ת��Ϊ�ַ�������
	 * @param field
	 * @param value
	 * @return
	 */
	private String formatObjValueToStr(Field field,Object value) {
		String strValue ="";
		if(value!=null) {
			if(Date.class.equals(field.getType())) {
				SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd hh:mm:ss");
				strValue = dateFormat.format(value);
			}else {
				//��ȡһ����������
				strValue = String.valueOf(value);
			}
		}
		return strValue;
	}
	
	/**
	 * �ж��Ƿ���ָ�����͵��ļ�
	 * @param file �ļ�����
	 * @param suff ��׺�����㿪ͷ�����磺.xls��
	 * @return
	 */
	public boolean isRightFile(File file,String suff) {
		return getFileSuffix(file).equals(suff);
	}
	
	/**
	 * @param file
	 * @return ����ָ���ļ���׺��
	 */
	public String getFileSuffix(File file) {
		String fileName = file.getName();
		int index =fileName.lastIndexOf('.');
		return fileName.substring(index, fileName.length());
	}

	/**
	 * ����ָ��ע��ĸ���
	 * @return
	 */
	public int getExcelAnnoCount() {
		return excelAnnoCount;
	}
}
