package com.excel.common_excel.service;

import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;

public interface ExcelService {

	<T> Workbook getExcelSheet(final List<T> data, final Class<T> clazz, final String sheetName,
			final String workbookName);

	<T> ResponseEntity<Resource> getExcelSheetAsResource(final List<T> data, final Class<T> clazz,
			final String sheetName, final String workbookName);

}
