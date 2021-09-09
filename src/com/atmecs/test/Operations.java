package com.atmecs.test;

import java.io.IOException;

import org.testng.annotations.Test;

import com.atmecs.utility.ExcelOperation;

public class Operations {

	@Test
	public void testUtil() throws IOException{
		ExcelOperation exo = new ExcelOperation();
		exo.isFilePresent("DataW");
		exo.openFileIfExist("DataWB");
		exo.getRowCount("DataWB");
		exo.getColumnCount("DataWB");
		exo.getCellNumber("DataWB", "Paper");
		exo.deleteRow("DataWB", 7);
//		exo.getRowCount("DataWB");
		exo.deleteColumnIfBlank("DataWB");
		exo.writeToCell("DataWB");
	}
}
