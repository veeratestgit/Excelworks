package Excelworks;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Readingdata {

//hirarachey
	//Excel--->Workbook----->Sheet---->Rows----->Columns
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		//first open excel
		FileInputStream file = new FileInputStream(System.getProperty("user.dir")+"\\data\\testdata.xlsx");
		
		//to get workbook
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		//now sheet anyone cannbe used
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		//XSSFSheet sheet =workbook.getsheetAt(0)
		//find numner of rows and colums
		
		int totalrows=sheet.getLastRowNum();
		//first get one row and then find cell in that
		int totalcells=sheet.getRow(1).getLastCellNum();
		
		System.out.println("Numner of rows:"+totalrows);
		System.out.println("Numner of columns:"+totalcells);
	
		for(int r=0;r<=totalrows;r++)
		{
			XSSFRow currentRow = sheet.getRow(r);
			
			for(int c=0;c<totalcells;c++)
			{
				XSSFCell cell = currentRow.getCell(c);
				System.out.print(cell.toString()+"\t");
			}
			System.out.println();
		}
		workbook.close();
		file.close();
	
	
	}

}
