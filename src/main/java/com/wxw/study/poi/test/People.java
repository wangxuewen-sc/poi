package com.wxw.study.poi.test;

import java.util.Date;

import com.wxw.util.excelIo.DataType;
import com.wxw.util.excelIo.ExcelProperty;

public class People {
	@ExcelProperty(zh="����",sequence=3)
	private Date birthday;
	@ExcelProperty(zh="����",sequence=0)
	private String name;
	@ExcelProperty(zh="�ֻ�����",sequence=2,dataType=DataType.STRING)
	private String phoneNum;
	@ExcelProperty(zh="����",sequence=1)
	private int age;
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public int getAge() {
		return age;
	}
	public void setAge(int age) {
		this.age = age;
	}
	public String getPhoneNum() {
		return phoneNum;
	}
	public void setPhoneNum(String phoneNum) {
		this.phoneNum = phoneNum;
	}
	public Date getBirthday() {
		return birthday;
	}
	public void setBirthday(Date birthday) {
		this.birthday = birthday;
	}
	@Override
	public String toString() {
		return "People [name=" + name + ", age=" + age + ", phoneNum=" + phoneNum + ", birthday=" + birthday + "]";
	}
	
	
}
