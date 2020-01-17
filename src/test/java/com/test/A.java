package com.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class A {

	public static String getDate(int rowNo, int cellno) throws Throwable  {
		String v = null;
		File loc = new File("C:\\Users\\Vignesh Chinnappa\\eclipse-workspace\\Integration\\Testdata\\file_example_XLSX_10.xlsx");
		FileInputStream stream = new FileInputStream(loc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet("Sheet1");
		Row r = s.getRow(rowNo);
		Cell c = r.getCell(cellno);
		int type = c.getCellType();
			if(type==1) {
				v = c.getStringCellValue();
			}
			else if(type==0) {
				if(DateUtil.isCellDateFormatted(c)) {
					Date dateCellValue = c.getDateCellValue();
					SimpleDateFormat sim = new SimpleDateFormat("dd-MMM-yyyy");
					v = sim.format(dateCellValue);			
				}
				else {
				double numericCellValue = c.getNumericCellValue();
				long l=(long)numericCellValue;
				v = String.valueOf(l);
				}
			}
		return v;
		
	}

}
