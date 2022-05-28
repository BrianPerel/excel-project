package com.excel.loader;

public class MyData {

  private String name;
  private String date;
  
  public MyData(String name, String date) {
    this.name = name;
    this.date = date;
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
    return getName() + " " + getDate();
  }
  
}
