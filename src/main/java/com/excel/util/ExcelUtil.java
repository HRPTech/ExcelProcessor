package com.excel.util;

import java.io.ByteArrayOutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

import org.apache.commons.lang.StringEscapeUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ByteArrayResource;

import com.excel.common_excel.config.ColumnConfig;

public final class ExcelUtil {

	private ExcelUtil() {
	}

	public static SXSSFWorkbook createWorkbook() {
		XSSFWorkbook workbook = new XSSFWorkbook();
		return new SXSSFWorkbook(workbook);
	}

	public static Map<String, SXSSFSheet> createSheets(final SXSSFWorkbook workbook, final List<String> sheetNames) {
		Map<String, SXSSFSheet> sheets = new HashMap<>();
		sheetNames.forEach(sheetName -> sheets.put(sheetName, workbook.createSheet(sheetName)));
		return sheets;
	}

	public static void createHeaderCells(final SXSSFSheet sheet, final List<ColumnConfig> columnConfig) {
		sheet.createRow(0);

		AtomicInteger atomicInteger = new AtomicInteger();
		Row headerRow = sheet.getRow(0);

		columnConfig.forEach(config -> {
			Cell cell = headerRow.createCell(atomicInteger.get());
			cell.setCellValue(config.getHeader());
			sheet.setColumnWidth(atomicInteger.get(), config.getColumnWidth());
			config.setCellIndex(atomicInteger.get());
			atomicInteger.getAndIncrement();
		});
		sheet.createRow(1);
	}

	public static ByteArrayResource writeWorkbookContents(final SXSSFWorkbook workbook) {
		ByteArrayOutputStream data = new ByteArrayOutputStream();

		try {
			workbook.write(data);
		} catch (Exception e) {
			throw new RuntimeException("error writing excel", e);
		}

		return new ByteArrayResource(data.toByteArray());
	}
	
	public static String asExcelFilename(final String fileName){
		return String.format("attachment; filename=%s.xlsx", StringEscapeUtils.unescapeXml(fileName));
	}

}
