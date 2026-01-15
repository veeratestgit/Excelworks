package Excelworks;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingDynamicdata {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		FileOutputStream file = new FileOutputStream(System.getProperty("user.dir")+"\\data\\writedynamicdata.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		
		XSSFSheet sheet = workbook.createSheet("DynamicData");
		
		Scanner sc = new Scanner(System.in);
		System.out.println("Enter how many rows");
		int rows = sc.nextInt();
		System.out.println("Enter how many cells");
		int cells = sc.nextInt();
		
		for(int r=0;r<=rows;r++)
		{
			XSSFRow currentrow =sheet.createRow(r);
			for(int c=0;c<cells;c++)
			{
				XSSFCell currentcell =currentrow.createCell(c);
				currentcell.setCellValue(sc.next());//used sc.next bcz user can enter any value
			}
		}
		workbook.write(file); //attach workbook to file
		workbook.close();
		file.close();
		System.out.println("File is created......");

	}

}
