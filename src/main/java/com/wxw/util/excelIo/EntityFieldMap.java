package com.wxw.util.excelIo;

public class EntityFieldMap {
	private String zh;//中文名
	private String filedName;//属性名
	private int sequence;//排序序列
	private boolean requisite;
	
	public EntityFieldMap() {}
	public EntityFieldMap(String filedName,String zh, int sequence,boolean requisite) {
		super();
		this.zh = zh;
		this.filedName = filedName;
		this.sequence = sequence;
		this.requisite = requisite;
	}
	
	public String getZh() {
		return zh;
	}
	public void setZh(String zh) {
		this.zh = zh;
	}
	public String getFiledName() {
		return filedName;
	}
	public void setFiledName(String filedName) {
		this.filedName = filedName;
	}
	public int getSequence() {
		return sequence;
	}
	public void setSequence(int sequence) {
		this.sequence = sequence;
	}
	public boolean isRequisite() {
		return requisite;
	}
	public void setRequisite(boolean requisite) {
		this.requisite = requisite;
	}
	@Override
	public String toString() {
		return "EntityFieldMap [zh=" + zh + ", filedName=" + filedName + ", sequence=" + sequence+ ", requisite=" + requisite + "]";
	}
	
	
}
