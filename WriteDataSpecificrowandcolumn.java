package Excelworks;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteDataSpecificrowandcolumn {

	public static void main(String[] args) throws FileNotFoundException {
		// TODO Auto-generated method stub
		FileOutputStream file = new FileOutputStream(System.getProperty("user.dir")+"\\data\\writerandomdata.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Data");
		XSSFRow row =sheet.createRow(4);
		XSSFCell cell = row.createCell(3);
		cell.setCellValue("Welcome");

	}

}
