package com.excel.common_excel;

import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;

import com.excel.common_excel.service.ExcelService;
import com.excel.common_excel.service.ExcelServiceImpl;

public class TestExcelDownload {
	
	private ExcelService service = new ExcelServiceImpl();

	@Test
	public void test() {
		Person hiran = new Person("Hiran", 1);
		Person sheetal = new Person("Sheetal", 2);
		
		List<Person> personList = new ArrayList<>();
		personList.add(hiran);
		personList.add(sheetal);
		
		ResponseEntity<Resource> resource = service.getExcelSheetAsResource(personList, Person.class, "Person sheet", "person workbook");
		

	}

}
