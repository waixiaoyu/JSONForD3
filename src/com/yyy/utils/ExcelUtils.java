package com.yyy.utils;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {

	private static Workbook parseFileName(String fileName) throws IOException {
		InputStream stream = new FileInputStream(fileName);
		String fileType = fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());
		Workbook wb = null;
		if (fileType.equals("xls")) {
			wb = new HSSFWorkbook(stream);
		} else if (fileType.equals("xlsx")) {
			wb = new XSSFWorkbook(stream);
		} else {
			System.out.println("您输入的excel格式不正确,退出！");
			System.exit(1);
		}
		return wb;
	}

	public static List<Row> getRows(String fileName, int nSheetIndex) throws IOException {
		Workbook wb = parseFileName(fileName);
		List<Row> lRows = new ArrayList<>();
		Sheet sheet = wb.getSheetAt(nSheetIndex);
		for (int i = 0; i < sheet.getLastRowNum(); i++) {
			lRows.add(sheet.getRow(i));
		}
		return lRows;
	}
}
