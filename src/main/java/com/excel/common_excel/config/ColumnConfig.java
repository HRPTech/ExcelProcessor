package com.excel.common_excel.config;

import com.excel.helper.SpreadsheetCellHelper.Style;

public class ColumnConfig implements Comparable<ColumnConfig> {

	private String header;
	private int columnWidth;
	private String methodName;
	private int cellIndex;
	private Style style;

	public ColumnConfig(String header, int columnWidth, String methodName, int cellIndex, Style style) {
		super();
		this.header = header;
		this.columnWidth = columnWidth;
		this.methodName = methodName;
		this.cellIndex = cellIndex;
		this.style = style;

	}

	public String getHeader() {
		return header;
	}

	public void setHeader(String header) {
		this.header = header;
	}

	public int getColumnWidth() {
		return columnWidth;
	}

	public void setColumnWidth(int columnWidth) {
		this.columnWidth = columnWidth;
	}

	public String getMethodName() {
		return methodName;
	}

	public void setMethodName(String methodName) {
		this.methodName = methodName;
	}

	public int getCellIndex() {
		return cellIndex;
	}

	public void setCellIndex(int cellIndex) {
		this.cellIndex = cellIndex;
	}

	public Style getStyle() {
		return style;
	}

	@Override
	public int compareTo(ColumnConfig o) {
		return Integer.compare(cellIndex, o.getCellIndex());
	}

}
