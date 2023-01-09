package org.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDrivenFramework {

	public static void main(String[] args) throws IOException {

		File excelLoc = new File("C:\\Users\\salin\\eclipse-workspace\\DayTwoMaven\\Excel\\Data.xlsx");

		FileInputStream fin = new FileInputStream(excelLoc);

		Workbook workbook = new XSSFWorkbook(fin);

		Sheet sheet = workbook.getSheet("Sheet1");

		/*
		  Row row = sheet.getRow(0); 
		  System.out.println(row); 
		 
		  Cell cell =row.getCell(0); 
		  System.out.println(cell);
		  
		  int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		  System.out.println(physicalNumberOfRows);
		  
		  Row row = sheet.getRow(0);
		  
		  int physicalNumberOfCells = row.getPhysicalNumberOfCells();
		  System.out.println(physicalNumberOfCells);
		 */

		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			
			Row row = sheet.getRow(i);

			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				
				Cell cell = row.getCell(j);
				//System.out.println(cell);
				
				CellType cellType = cell.getCellType();
				
				switch (cellType) {
				case STRING:
					// to get the String value from the cell
					String stringCellValue = cell.getStringCellValue();
					System.out.println(stringCellValue);
					break;
				
				case NUMERIC: // In numeric we have date also we have to split
					
					// to get the date value from the cell
					if (DateUtil.isCellDateFormatted(cell)) // DateUtil => class 
					{	
						// isCellDateFormatted static method returns boolean
						Date dateCellValue = cell.getDateCellValue();
						// Date is a class 
						
						// need to assign the date format
						
						SimpleDateFormat dateForm = new SimpleDateFormat("dd-MM-yyyy");
						// SimpleDateformat is a class
						
						String format = dateForm.format(dateCellValue);
						System.out.println(format);
					}
					else // to get the numeric value from the cell (eg:phone number) 
					{
						double numericCellValue = cell.getNumericCellValue();
						// conert to long
						long l = (long) numericCellValue;
						System.out.println(l);
						break;
					}
				
				default:
					break;
				}

			}
		}
	}
}