package com.excel.helper;

import java.math.BigDecimal;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

import com.excel.common_excel.ExcelConstants;

public final class SpreadsheetCellHelper {

	private SpreadsheetCellHelper() {
	}

	public enum Style {
		NO_STYLE {
			CellStyle apply(final Workbook workbook) {
				final CellStyle style = workbook.createCellStyle();
				return style;
			}
		},
		TITLE {
			CellStyle apply(final Workbook workbook) {
				final CellStyle style = workbook.createCellStyle();
				style.setFillPattern(FillPatternType.SPARSE_DOTS);
				return style;
			}
		},
		TWO_NUMBER_FORMAT {
			CellStyle apply(final Workbook workbook) {
				final CellStyle style = workbook.createCellStyle();
				style.setAlignment(HorizontalAlignment.RIGHT);
				style.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(ExcelConstants.TWO_NUMBER_FORMAT));
				return style;
			}
		};

		abstract CellStyle apply(final Workbook workbook);
	}

	public static void setValue(final Cell cell, final Object value) {
		if (value instanceof String) {
			setValue(cell, (String) value);
		} else if (value instanceof Integer) {
			setValue(cell, (Integer) value);
		} else if (value instanceof Enum) {
			setValue(cell, value.toString());
		} else if (value instanceof BigDecimal) {
			setValue(cell, ((BigDecimal) value).doubleValue());
		}
	}

	public static void applyStyle(final Workbook workbook, final Cell cell, final Style style) {
		cell.setCellStyle(style.apply(workbook));
	}

	/*
	 * Can apply any business logic if required
	 */
	private static void setValue(final Cell cell, final String value) {
		cell.setCellValue(value);
	}

	private static void setValue(final Cell cell, final Integer value) {
		cell.setCellValue(value);
	}
	
	private static void setValue(final Cell cell, final Double value) {
		cell.setCellValue(value);
	}

}
