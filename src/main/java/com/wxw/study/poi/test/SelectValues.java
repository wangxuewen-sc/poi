package com.wxw.study.poi.test;

import java.util.ArrayList;
import java.util.List;

public class SelectValues {
	
	public List<String> findDepartment(){
		List<String> departList = new ArrayList<>();
		departList.add("技术部");
		departList.add("市场部");
		departList.add("销售部");
		departList.add("维护部");
		return departList;
	}
	
	public List<String> findPosition(){
		List<String> positionList = new ArrayList<>();
		positionList.add("市场部总监");
		positionList.add("市场部经理");
		positionList.add("销售部经理");
		positionList.add("CEO");
		return positionList;
	}
	
	public List<String> findProvinceCascade(){
		
		List<String> shengShi = new ArrayList<>();
		shengShi.add("四川");
		shengShi.add("云南");
		shengShi.add("江苏");
		return shengShi;
	}
	
public List<String> findCities(){
		
		List<String> cities = new ArrayList<>();
		cities.add("成都");
		cities.add("自贡");
		cities.add("广元");
		cities.add("绵阳");
		cities.add("德阳");
		cities.add("景洪");
		cities.add("昆明");
		cities.add("曲靖");
		cities.add("玉溪");
		cities.add("宝山");
		cities.add("南京");
		cities.add("苏州");
		cities.add("无锡");
		cities.add("常州");
		cities.add("南通");
		cities.add("泰州");
		cities.add("扬州");
		cities.add("盐城");
		cities.add("镇江");
		cities.add("宿迁");
		cities.add("淮安");
		cities.add("徐州");
		cities.add("连云港");
		return cities;
	}
}
