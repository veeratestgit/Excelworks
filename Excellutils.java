package Excelworks; // Package name to organize the class

// Import required Java classes for file handling
import java.io.FileInputStream;     // To read data from an Excel file
import java.io.FileOutputStream;    // To write data back into an Excel file
import java.io.IOException;         // To handle input-output exceptions

// Import Apache POI classes for Excel operations
import org.apache.poi.ss.usermodel.CellStyle;        // To apply styles (like colors) to cells
import org.apache.poi.ss.usermodel.DataFormatter;    // To format cell values into human-readable strings
import org.apache.poi.ss.usermodel.FillPatternType;  // To define how the color/pattern should fill the cell
import org.apache.poi.ss.usermodel.IndexedColors;    // Predefined colors (Red, Green, etc.)
import org.apache.poi.xssf.usermodel.XSSFCell;       // Represents a cell in Excel (.xlsx)
import org.apache.poi.xssf.usermodel.XSSFRow;        // Represents a row in Excel
import org.apache.poi.xssf.usermodel.XSSFSheet;      // Represents a sheet in Excel
import org.apache.poi.xssf.usermodel.XSSFWorkbook;   // Represents the entire Excel workbook (.xlsx)

// Utility class to work with Excel files
public class Excellutils {
	
	// Declare reusable static variables
	public static FileInputStream fi;   // Reads Excel file
	public static FileOutputStream fo;  // Writes Excel file
	public static XSSFWorkbook wb;      // Workbook object
	public static XSSFSheet ws;         // Sheet object
	public static XSSFRow row;          // Row object
	public static XSSFCell cell;        // Cell object
	public static CellStyle style;      // Style object (for formatting cells)

	// ------------------------ METHOD 1 ------------------------
	// Method to get row count from a given sheet
	public static int getRowCount(String xlfile, String xlsheet) throws IOException
	{
		fi = new FileInputStream(xlfile);     // Open Excel file in read mode
		wb = new XSSFWorkbook(fi);            // Load workbook
		ws = wb.getSheet(xlsheet);            // Get sheet by name
		int rowcount = ws.getLastRowNum();    // Get index of last row (0-based)
		wb.close();                           // Close workbook to free memory
		fi.close();                           // Close file stream
		return rowcount;                      // Return row count
	}

	// ------------------------ METHOD 2 ------------------------
	// Method to get cell count from a specific row
	public static int getCellCount(String xlfile, String xlsheet, int rownum) throws IOException
	{
		fi = new FileInputStream(xlfile);     
		wb = new XSSFWorkbook(fi);            
		ws = wb.getSheet(xlsheet);            
		row = ws.getRow(rownum);              // Get the row object
		int cellcount = row.getLastCellNum(); // Get number of cells in that row
		wb.close();                           
		fi.close();                           
		return cellcount;                     
	}

	// ------------------------ METHOD 3 ------------------------
	// Method to read cell data from a given row and column
	public static String getCellData(String xlfile, String xlsheet, int rownum, int column) throws IOException
	{
		fi = new FileInputStream(xlfile);     
		wb = new XSSFWorkbook(fi);            
		ws = wb.getSheet(xlsheet);            
		row = ws.getRow(rownum);              // Get row
		cell = row.getCell(column);           // Get cell from row
		String data;
		try
		{
			// DataFormatter ensures data is read properly irrespective of type (string, number, date, formula)
			DataFormatter formatter = new DataFormatter();
			data = formatter.formatCellValue(cell); 
		}
		catch(Exception e)
		{
			data = ""; // If any error (e.g., null cell), return empty string
		}
		wb.close(); 
		fi.close();
		return data; // Return the extracted data
	}

	// ------------------------ METHOD 4 ------------------------
	// Method to write data into a cell
	public static void SetCellData(String xlfile, String xlsheet, int rownum, int column, String data) throws IOException
	{
		fi = new FileInputStream(xlfile);     
		wb = new XSSFWorkbook(fi);            
		ws = wb.getSheet(xlsheet);            
		row = ws.getRow(rownum);              
		cell = row.createCell(column);        // Create a new cell if it doesn’t exist
		cell.setCellValue(data);              // Set the provided data into the cell
		fo = new FileOutputStream(xlfile);    // Open Excel file in write mode
		wb.write(fo);                         // Save changes to workbook
		wb.close(); 
		fi.close(); 
		fo.close(); 
	}

	// ------------------------ METHOD 5 ------------------------
	// Method to fill a cell with GREEN color (e.g., for "PASS" status)
	public static void fillGreenColor(String xlfile, String xlsheet, int rownum, int column) throws IOException
	{
		fi = new FileInputStream(xlfile);     
		wb = new XSSFWorkbook(fi);            
		ws = wb.getSheet(xlsheet);            
		row = ws.getRow(rownum);              
		cell = row.getCell(column);           
		
		style = wb.createCellStyle();          // Create a style object
		style.setFillForegroundColor(IndexedColors.GREEN.getIndex()); // Set foreground color = Green
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);       // Fill style = solid color
		
		cell.setCellStyle(style);             // Apply style to the cell
		fo = new FileOutputStream(xlfile);    
		wb.write(fo);                         
		wb.close(); 
		fi.close(); 
		fo.close(); 
	}

	// ------------------------ METHOD 6 ------------------------
	// Method to fill a cell with RED color (e.g., for "FAIL" status)
	public static void fillRedColor(String xlfile, String xlsheet, int rownum, int column) throws IOException
	{
		fi = new FileInputStream(xlfile);     
		wb = new XSSFWorkbook(fi);            
		ws = wb.getSheet(xlsheet);            
		row = ws.getRow(rownum);              
		cell = row.getCell(column);           
		
		style = wb.createCellStyle();          
		style.setFillForegroundColor(IndexedColors.RED.getIndex());   // Set foreground color = Red
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);       // Solid fill style
		
		cell.setCellStyle(style);             
		fo = new FileOutputStream(xlfile);    
		wb.write(fo);                         
		wb.close(); 
		fi.close(); 
		fo.close(); 
	}
}

