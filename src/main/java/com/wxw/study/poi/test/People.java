package com.wxw.study.poi.test;

import java.util.Date;

import com.wxw.util.excelIo.DataType;
import com.wxw.util.excelIo.ExcelProperty;
import com.wxw.util.excelIo.Select;

public class People {
	@ExcelProperty(zh="生日",sequence=3)
	private Date birthday;
	@ExcelProperty(zh="姓名",sequence=0)
	private String name;
	@ExcelProperty(zh="手机号码",sequence=2,dataType=DataType.STRING)
	private String phoneNum;
	@ExcelProperty(zh="年龄",sequence=1)
	private int age;
	@ExcelProperty(zh="部门",sequence=4)
	@Select(method="findDepartment",clazz=SelectValues.class)
	private String department ;
	@ExcelProperty(zh="职位",sequence=5)
	@Select(method="findPosition",clazz=SelectValues.class)
	private String position;
	@ExcelProperty(zh="省",sequence=6)
	@Select(method="findProvinceCascade",clazz=SelectValues.class)
	private String province;
	@ExcelProperty(zh="市",sequence=7)
	@Select(method="findCities",clazz=SelectValues.class)
	private String city;
	
	
	public Date getBirthday() {
		return birthday;
	}
	public void setBirthday(Date birthday) {
		this.birthday = birthday;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getPhoneNum() {
		return phoneNum;
	}
	public void setPhoneNum(String phoneNum) {
		this.phoneNum = phoneNum;
	}
	public int getAge() {
		return age;
	}
	public void setAge(int age) {
		this.age = age;
	}
	public String getDepartment() {
		return department;
	}
	public void setDepartment(String department) {
		this.department = department;
	}
	public String getPosition() {
		return position;
	}
	public void setPosition(String position) {
		this.position = position;
	}
	
}
