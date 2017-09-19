package com.excel.common_excel.service;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;

import com.excel.common_excel.annotation.ExcelCell;
import com.excel.common_excel.config.ColumnConfig;
import com.excel.helper.SpreadsheetCellHelper;
import com.excel.util.ExcelUtil;

@Service
public class ExcelServiceImpl implements ExcelService {

	private <T> List<ColumnConfig> buildExcelConfig(final Class<T> clazz) {

		List<ColumnConfig> config = new ArrayList<>();
		Field[] fields = clazz.getDeclaredFields();
		for (Field field : fields) {
			if (field.isAnnotationPresent(ExcelCell.class)) {
				ExcelCell cell = field.getAnnotationsByType(ExcelCell.class)[0];
				ColumnConfig columnConfig = new ColumnConfig(cell.header(), cell.width(), field.getName(), cell.order(),
						cell.style());
				config.add(columnConfig);
			}
		}
		return config.stream().sorted().collect(Collectors.toList());
	}

	private SXSSFWorkbook multiSheetWorkbook(final List<List<ColumnConfig>> excelConfig, final List<String> sheetNames,
			final SXSSFWorkbook workbook) {

		if (excelConfig.size() != sheetNames.size()) {
			throw new RuntimeException("The excel config size and sheet size are different");
		}

		SXSSFWorkbook sxssfWorkbook = workbook;
		if (workbook == null) {
			sxssfWorkbook = ExcelUtil.createWorkbook();
		}

		Map<String, SXSSFSheet> sheets = ExcelUtil.createSheets(sxssfWorkbook, sheetNames);

		for (int i = 0; i < excelConfig.size(); i++) {
			ExcelUtil.createHeaderCells(sheets.get(sheetNames.get(i)), excelConfig.get(i));
		}

		return sxssfWorkbook;
	}

	private <T> void writeContentToSheetUsingAnnotation(final SXSSFWorkbook workbook, final SXSSFSheet sheet,
			final List<T> data, final Class<T> clazz, final List<ColumnConfig> excelConfig,
			final String localeFormatStyle) {
		Class<? extends T> localClass = clazz;
		if (sheet == null) {
			throw new RuntimeException("The sheet must exist in the workbook");
		}

		AtomicInteger rowNumber = new AtomicInteger(1);
		if (!CollectionUtils.isEmpty(data)) {
			data.forEach(d -> {
				final Row row = sheet.createRow(rowNumber.getAndIncrement());
				AtomicInteger cellNumber = new AtomicInteger(0);

				excelConfig.forEach(con -> {
					Cell cell = row.createCell(cellNumber.getAndIncrement());
					try {
						// set the cell value
						Optional<Object> optionalValue = Optional
								.ofNullable(PropertyUtils.getProperty(localClass.cast(d), con.getMethodName()));
						optionalValue.ifPresent(value -> {
							SpreadsheetCellHelper.setValue(cell, value);
						});

						// formatting
						SpreadsheetCellHelper.applyStyle(workbook, cell, con.getStyle());

					} catch (Exception e) {
						throw new RuntimeException("error writing excel", e);
					}
				});
			});
		}
	}

	public SXSSFWorkbook generateWorkbook() {
		return ExcelUtil.createWorkbook();
	}

	public <T> void generateMultiSheetExcel(final List<T> data, final Class<T> clazz, final String sheetName,
			final Workbook book) {
		SXSSFWorkbook workbook = (SXSSFWorkbook) book;
		if (book == null) {
			workbook = ExcelUtil.createWorkbook();
		}
		generateExcel(data, clazz, sheetName, workbook);
	}

	public <T> ResponseEntity<Resource> getSingleExcelSheetAsResource(final List<T> data, final Class<T> clazz,
			final String sheetName, final String workbookName) {
		SXSSFWorkbook workbook = generateExcel(data, clazz, sheetName, null);
		ByteArrayResource resource = ExcelUtil.writeWorkbookContents(workbook);
		workbook.dispose();
		return ok(resource, workbookName);
	}

	public <T> ResponseEntity<Resource> generateExcelResource(final Workbook workbook, final String name) {
		ByteArrayResource resource = ExcelUtil.writeWorkbookContents(workbook);
		return ok(resource, name);
	}

	private <T> SXSSFWorkbook generateExcel(final List<T> data, final Class<T> clazz, final String sheetName,
			final SXSSFWorkbook book) {
		List<ColumnConfig> columnConfig = buildExcelConfig(clazz);

		SXSSFWorkbook workbook = multiSheetWorkbook(Collections.singletonList(columnConfig),
				Collections.singletonList(sheetName), book);

		writeContentToSheetUsingAnnotation(workbook, workbook.getSheet(sheetName), data, clazz, columnConfig, null);

		return workbook;
	}

	private ResponseEntity<Resource> ok(final ByteArrayResource resource, final String name) {
		return ResponseEntity.ok().header("Content-Disposition", ExcelUtil.asExcelFilename(name))
				.header("Content-Encoding", "UTF-8").contentLength(resource.contentLength())
				.contentType(
						MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
				.body(resource);
	}
}
