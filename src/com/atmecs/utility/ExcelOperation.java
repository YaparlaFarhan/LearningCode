package com.atmecs.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOperation {
	public void openFileIfExist(String fileName) throws IOException {
		String filepath = "C:\\Users\\farhan.yaparla\\Desktop\\Testing\\" + fileName + ".xlsx";
		FileInputStream infile = new FileInputStream(filepath);

		XSSFWorkbook workb = new XSSFWorkbook(infile);
		System.out.println(fileName + " file has been open");
		infile.close();

	}

	public void isFilePresent(String fileName) {
		String filepath = "C:\\Users\\farhan.yaparla\\Desktop\\Testing\\" + fileName + ".xlsx";
		try {
			FileInputStream infile = new FileInputStream(filepath);
			System.out.println(fileName + "is present");
		} catch (FileNotFoundException e) {

			System.out.println(fileName + " is not present");
		}

	}

	public void getRowCount(String fileName) throws IOException {
		String filepath = "C:\\Users\\farhan.yaparla\\Desktop\\Testing\\" + fileName + ".xlsx";
		FileInputStream infile = new FileInputStream(filepath);

		XSSFWorkbook workb = new XSSFWorkbook(infile);
		XSSFSheet wbsheet = workb.getSheetAt(0);

		int rows = wbsheet.getLastRowNum();
		System.out.println("No 0f rows" + rows);
		infile.close();
	}

	public void getColumnCount(String fileName) throws IOException {
		String filepath = "C:\\Users\\farhan.yaparla\\Desktop\\Testing\\" + fileName + ".xlsx";
		FileInputStream infile = new FileInputStream(filepath);

		XSSFWorkbook workb = new XSSFWorkbook(infile);
		XSSFSheet wbsheet = workb.getSheetAt(0);

		int rows = wbsheet.getLastRowNum();
		int cols = wbsheet.getRow(1).getLastCellNum();

		System.out.println("No of coloms" + cols);
		infile.close();

	}

	public void getCellNumber(String fileName, String data) throws IOException {
		String filepath = "C:\\Users\\farhan.yaparla\\Desktop\\Testing\\" + fileName + ".xlsx";
		FileInputStream infile = new FileInputStream(filepath);

		XSSFWorkbook workb = new XSSFWorkbook(infile);
		XSSFSheet wbsheet = workb.getSheetAt(0);

		int rows = wbsheet.getLastRowNum();
		int cols = wbsheet.getRow(1).getLastCellNum();
		for (int r = 1; r <= rows; r++) {
			XSSFRow row = wbsheet.getRow(r);

			for (int c = 0; c < cols; c++) {
				XSSFCell cell = row.getCell(c);
				switch (cell.getCellType()) {
				case STRING:
//					System.out.print(cell.getStringCellValue() + "    ");
					if (data.equals(cell.getStringCellValue())) {
						System.out.println("Cell Number for a " + data + "is " + r + "," + c);
						break;

					} else {

					}
					break;
				case NUMERIC:
//					System.out.print(cell.getNumericCellValue() + "    ");
					if (data.equals(cell.getNumericCellValue())) {
						System.out.println(data + " cell and row num " + r + "," + c);
						break;

					} else {

					}

					break;
				default:

					break;

				}
			}
		}
		infile.close();

	}

//	public void deleteRow(String fileName) throws IOException {
//		String filepath = "C:\\Users\\farhan.yaparla\\Desktop\\Testing\\" + fileName + ".xlsx";
//		FileInputStream infile = new FileInputStream(filepath);
//
//		XSSFWorkbook workb = new XSSFWorkbook(infile);
//		XSSFSheet wbsheet = workb.getSheetAt(0);
//
//		int rows = wbsheet.getLastRowNum();
//		int cols = wbsheet.getRow(1).getLastCellNum();
//		wbsheet.removeRowBreak(7);
//		wbsheet.removeRow(wbsheet.getRow(8));
//		infile.close();
//	}

	public void deleteRow(String fileName, int rowNo) throws IOException {
		String filepath = "C:\\Users\\farhan.yaparla\\Desktop\\Testing\\" + fileName + ".xlsx";
		FileInputStream infile = new FileInputStream(filepath);

		XSSFWorkbook workb = new XSSFWorkbook(infile);
		XSSFSheet wbsheet = workb.getSheetAt(0);
//		int rowbefore=wbsheet.getLastRowNum();

		int lastRowNum = wbsheet.getLastRowNum();
		if (rowNo >= 0 && rowNo < lastRowNum) {
			wbsheet.shiftRows(rowNo, lastRowNum, -1);
		}
		if (rowNo == lastRowNum) {
			XSSFRow removingrow = wbsheet.getRow(rowNo);
			if (removingrow != null) {
				wbsheet.removeRow(removingrow);
			}
		}
		infile.close();
		FileOutputStream outfile = new FileOutputStream(new File(filepath));
		workb.write(outfile);
		int rowsafter = wbsheet.getLastRowNum();
		outfile.close();

		System.out.println(lastRowNum + " " + rowsafter + " Row has deleted sucessfully");

	}

	public void deleteColumnIfBlank(String fileName) throws IOException {
		String filepath = "C:\\Users\\farhan.yaparla\\Desktop\\Testing\\" + fileName + ".xlsx";
		FileInputStream infile = new FileInputStream(filepath);

		XSSFWorkbook workb = new XSSFWorkbook(infile);
		XSSFSheet wbsheet = workb.getSheetAt(0);

		int lastRowNum = wbsheet.getLastRowNum();
		int lastColNum = wbsheet.getRow(1).getLastCellNum();
		for (int r = 1; r <= lastRowNum; r++) {
			XSSFRow row = wbsheet.getRow(r);
			for (int c = 0; c < lastColNum; c++) {
				XSSFCell cell = row.getCell(c);
				if (cell.getStringCellValue().isBlank())
					cell.removeCellComment();
			}
		}
		infile.close();
		FileOutputStream outfile = new FileOutputStream(new File(filepath));
		workb.write(outfile);
		outfile.close();

		System.out.println("--");
	}

	public void writeToCell(String fileName) throws IOException {
		String filepath = "C:\\Users\\farhan.yaparla\\Desktop\\Testing\\" + fileName + ".xlsx";
		FileInputStream fins = new FileInputStream(filepath);
		XSSFWorkbook workbook = new XSSFWorkbook(fins);

		XSSFSheet sheet = workbook.getSheetAt(0);
		int lastRowNum = sheet.getLastRowNum();
		Object[][] data = { { "NE", "Test", "TestObject" } };
		int rows = data.length;
		int cols = data[0].length;

//		System.out.println(rows + " " + cols);
		for (int r = 0; r < rows; r++) {
			XSSFRow row = sheet.createRow(lastRowNum + 1);
			for (int c = 0; c < cols; c++) {
				XSSFCell cell = row.createCell(c);
				Object value = data[r][c];
				if (value instanceof String)
					cell.setCellValue((String) value);

			}
		}
		fins.close();
		FileOutputStream fis = new FileOutputStream(filepath);
		workbook.write(fis);
		fis.close();
		System.out.println("Data has writen sucessfully");

	}

//	public void writeToCell(String fileName, Object[][] data) throws IOException {
//		
//		System.out.println("writeToCell");
//
//		String filepath = "C:\\Users\\farhan.yaparla\\Desktop\\Testing\\" + fileName + ".xlsx";
//		FileInputStream infile = new FileInputStream(filepath);
//
//		XSSFWorkbook workb = new XSSFWorkbook();
//		XSSFSheet wbsheet = workb.getSheetAt(0);
//
//		int lastRowNum = wbsheet.getLastRowNum();
//
//		int rows = data.length;
//		int cols = data.length;
//
//		for (int r = lastRowNum + 1; r < lastRowNum + rows; r++) {
//			XSSFRow row = wbsheet.createRow(r);
//			for (int c = 0; c < cols; c++) {
//				XSSFCell cell = row.createCell(c);
//				Object value = data[0][c];
//				System.out.print(data[0][c]);
//				if(value instanceof String)
//				cell.setCellValue((String) value);
//			}
//		}
//		infile.close();
//		FileOutputStream fileOut = new FileOutputStream(filepath);
//		workb.write(fileOut);
//		fileOut.close();
//		System.out.println("Data has written in Excel");
//
//	}

//	public static void main(String[] args) throws IOException {
//
//		String filepath = ".\\Data\\DataWB.xlsx";
//		FileInputStream infile = new FileInputStream(filepath);
//
//		XSSFWorkbook workb = new XSSFWorkbook(infile);
//		XSSFSheet wbsheet = workb.getSheetAt(0);
//
//		int rows = wbsheet.getLastRowNum();
//		int cols = wbsheet.getRow(1).getLastCellNum();
//
//		for (int r = 0; r <= rows; r++) {
//			XSSFRow row = wbsheet.getRow(r);
//
//			for (int c = 0; c < cols; c++) {
//				XSSFCell cell = row.getCell(c);
//				switch (cell.getCellType()) {
//				case STRING:
//					System.out.print(cell.getStringCellValue() + "    ");
//					break;
//				case NUMERIC:
//					System.out.print(cell.getNumericCellValue() + "    ");
//					break;
//				default:
//
//					break;
//
//				}
//			}
//			System.out.println();
//		}
//
//	}
}
