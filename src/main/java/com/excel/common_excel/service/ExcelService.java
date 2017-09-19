package com.excel.common_excel.service;

import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;

public interface ExcelService {

	<T> void generateMultiSheetExcel(final List<T> data, final Class<T> clazz, final String sheetName,
			final Workbook book);

	<T> ResponseEntity<Resource> getSingleExcelSheetAsResource(final List<T> data, final Class<T> clazz,
			final String sheetName, final String workbookName);

	<T> ResponseEntity<Resource> generateExcelResource(final Workbook workbook, final String name);

	SXSSFWorkbook generateWorkbook();

}
