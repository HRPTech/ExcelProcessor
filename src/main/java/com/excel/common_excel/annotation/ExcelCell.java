package com.excel.common_excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import com.excel.helper.SpreadsheetCellHelper.Style;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelCell {
	
	String header();
	int order();
	int width() default 3000;
	Style style() default Style.NO_STYLE;
	
	
}
