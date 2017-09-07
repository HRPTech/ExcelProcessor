package com.excel.common_excel;

import com.excel.common_excel.annotation.ExcelCell;

public class Person {
	@ExcelCell(header="Deal Id", order=0)
	private String name;
	@ExcelCell(header="Deal Id", order=1)
	private int id;
	
	
	public Person(String name, int id) {
		super();
		this.name = name;
		this.id = id;
	}


	public String getName() {
		return name;
	}


	public int getId() {
		return id;
	}
	
	
	
}
