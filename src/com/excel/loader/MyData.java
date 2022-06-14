package com.excel.loader;

public class MyData {

	private String name;
	private String date;
	
	public MyData() {
		this("", "");
	}

	public MyData(String argName, String argDate) {
		this.name = argName;
		this.date = argDate;
	}

	public String getDate() {
		return date;
	}

	public String getName() {
		return name;
	}

	public void setDate(String argDate) {
		date = argDate;
	}

	public void setName(String argName) {
		name = argName;
	}

	@Override
	public String toString() {
		return String.format("%s %s", getName(), getDate());
	}
	
}
