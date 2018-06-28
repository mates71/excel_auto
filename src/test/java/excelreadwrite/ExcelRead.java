package excelreadwrite;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {
	
	public static void main(String[] args) throws IOException  {
		
		
		String excelPath= "./src/test/resources/TestData/EmpData.xlsx";
		//Open the file to read
				FileInputStream in=new FileInputStream(excelPath);
				//Let the Apache XSSFWorkbook class handle the data
				XSSFWorkbook workbook=new XSSFWorkbook(in);//xssf=spread sheet format
				//Jump to Worksheet level
				XSSFSheet worksheet=workbook.getSheet("Sheet1");
				//Find out how many rows
				int rowsCount=worksheet.getPhysicalNumberOfRows();
				
				System.out.println("Number of rows:" + rowsCount);
				//print first row and first cell data
				System.out.println(worksheet.getRow(0).getCell(0));
				//print second row and first cell data
				System.out.println(worksheet.getRow(1).getCell(0));
				
				String cellValue=worksheet.getRow(3).getCell(1).toString();
				System.out.println(cellValue);
				
				
			//	Print all Names
				int sheetRowsCount=worksheet.getPhysicalNumberOfRows();
				
				for(int row=1;row<sheetRowsCount;row++){
					String EmpID=worksheet.getRow(row).getCell(0).toString();
					String Name=worksheet.getRow(row).getCell(1).toString();
					
					String status=worksheet.getRow(row).getCell(2).getStringCellValue();
					System.out.println(EmpID+"-->"+Name+"-->"+status);	
					//Emp001-->John --> IT
				}
//			
				in.close();
			}

	}


