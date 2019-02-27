package com.wxw.study.poi.test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;

import com.wxw.util.excelIo.Excel;

public class PoiStudy {
	public static void main(String[] args) throws IOException, ParseException {
		File file = new File("E:/type.xlsx");
		FileOutputStream fos = new FileOutputStream(file);
		Excel<People> excel = new Excel<>(People.class);
		List<People> peoples = new ArrayList<People>();
		for(int i=0;i<1;i++) {
			People p =new People();
			p.setName("Íô"+i);
			p.setAge(20+i);
			p.setPhoneNum(null);
			try {
				p.setBirthday(new SimpleDateFormat("yyyy/MM/dd").parse("2018/10/31"));
			} catch (ParseException e) {
				e.printStackTrace();
			}
			peoples.add(p);
		}
		Workbook workbook = excel.createExcel(peoples);
		workbook.write(fos);
		
		
//		File file = new File("E:/type.xlsx");
//		Excel<People> excel = new Excel<>(People.class);
//		InputStream fis = new FileInputStream(file);
//		List<People> peoples = excel.readExcel(fis);
//		for(People p:peoples) {
//			System.out.println(p);
//			
//		}
		
		
//		SimpleDateFormat format = new SimpleDateFormat("EEE MMM dd HH:mm:ss Z yyyy", Locale.UK);
//		Date date = format.parse("Wed Oct 31 00:00:00 CST 2018");
//		format = new SimpleDateFormat("yyyy/MM/dd hh:mm:ss");
//		System.out.println(format.format(date));
		
	}
	
}
