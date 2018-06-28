package excelautomation_home;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

	public static void main(String[] args) throws IOException {
		
		String excelPath="/Users/musaates/Desktop/EmpData.xlsx";
//		//Open the file to read
//		FileInputStream input=new FileInputStream(excelPath);
//		//Let the Apache XSSFWorkbook class handle the data
//		XSSFWorkbook workbook=new XSSFWorkbook(input);
//		//Jump to Worksheet level
//		XSSFSheet worksheet=workbook.getSheet("Sheet1");
//		
//		//Find out how many rows
//		int countRows=worksheet.getPhysicalNumberOfRows();
//		System.out.println("Total rows : "+countRows);
//		
//		//print first row and first cell data
//		System.out.println("first row and first cell data :"+worksheet.getRow(0).getCell(0));
//		System.out.println("0,2 status "+worksheet.getRow(0).getCell(2));
//		System.out.println("Is it cucumber? "+worksheet.getRow(2).getCell(1));
//		System.out.println("which book? "+worksheet.getRow(4).getCell(1)); 
//		System.out.println("is ist android? "+worksheet.getRow(5).getCell(0));
//		//print second row and first cell data
//		System.out.println("2 R,1 C "+worksheet.getRow(2).getCell(1));
//
//		//Print all Names
//		int sheetRowsCount=worksheet.getPhysicalNumberOfRows();
//		for(int row=1;row<sheetRowsCount;row++){
//			String name=worksheet.getRow(row).getCell(1).toString();
//			String ID=worksheet.getRow(row).getCell(0).toString();
//			String status=worksheet.getRow(row).getCell(2).toString();
		FileInputStream in=new FileInputStream(excelPath);
		XSSFWorkbook wb=new XSSFWorkbook(in);
		XSSFSheet sh=wb.getSheet("Sheet2");
	
		XSSFRow row=sh.getRow(0);
		XSSFCell cell=row.getCell(0);
		int rowsCount=sh.getPhysicalNumberOfRows();
		int cellCount=row.getPhysicalNumberOfCells();
		
		System.out.println(sh.getRow(3).getCell(1));
		
		System.out.println(sh.getRow(3).getCell(2));
		System.out.println(sh.getRow(3).getCell(3));
		System.out.println("=========");
		
		int sheetRowsCount=sh.getPhysicalNumberOfRows();
		int cells=row.getPhysicalNumberOfCells();
		
		
		for (int i = 1; i < 5; i++) {
			
			for (int j = 1; j < 8; j++) {
				System.out.println(i+" "+j);
			}
		
		}
		
		
		
		
		
		
//		
//			String total=sh.getRow(i).getCell(1).toString();
//			
//			System.out.println(name);
//		}
////			
			//System.out.println(ID+"-->"+name+"-->"+status);
//			System.out.print(name);
//			System.out.print(ID);
		in.close();
		}
	
	}


