package Excelworks;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WrittingStaticdata {
	//Excel--->Workbook----->Sheet---->Rows----->Columns
	
	public static void main(String[] args) throws IOException {
		//open file in write mode
		FileOutputStream file = new FileOutputStream(System.getProperty("user.dir")+"\\data\\writedata.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		
		XSSFSheet sheet = workbook.createSheet("Data");
		//create row1 and add value into celss
		XSSFRow row1 =sheet.createRow(0);
		row1.createCell(0).setCellValue("Java");
		row1.createCell(1).setCellValue(12);
		row1.createCell(2).setCellValue("Automation");
		//create row2 and add value into celss
		XSSFRow row2 =sheet.createRow(0);
		row2.createCell(0).setCellValue("Python");
		row2.createCell(1).setCellValue(34);
		row2.createCell(2).setCellValue("Automation");
		//create row3 and add value into celss
		XSSFRow row3 =sheet.createRow(0);
		row3.createCell(0).setCellValue("C#");
		row3.createCell(1).setCellValue(13);
		row3.createCell(2).setCellValue("Automation");

		//NOW attach the workbook into file
		workbook.write(file);
		workbook.close();
		file.close();
		System.out.println("File is created");
	}

}
