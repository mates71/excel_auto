package excelautomation_home;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {

	public static void main(String[] args) throws IOException {

		String excelPath ="/Users/musaates/Desktop/EmpData.xlsx";
		
		FileInputStream in=new FileInputStream(excelPath);
		XSSFWorkbook wb=new XSSFWorkbook(in);
		XSSFSheet sh=wb.getSheet("Sheet2");
		XSSFRow row=sh.getRow(0);
		int rowsCount=sh.getPhysicalNumberOfRows();
		int cellCount=row.getPhysicalNumberOfCells();
		System.out.println(rowsCount);
		
		XSSFCell cell=sh.getRow(1).getCell(4);
		if(cell==null){
			cell=sh.getRow(1).createCell(4);
			
		}
		cell.setCellValue("DEMOOOO");
		
		
		cell=sh.getRow(3).getCell(3);
		if(cell==null){
			cell=sh.getRow(3).createCell(3);
		}
		cell.setCellValue("POSITIVE");
		
		
		cell=sh.getRow(5).getCell(5);
		if(cell==null){
			cell=sh.getRow(5).createCell(5);
		}
		cell.setCellValue("ALLAHIM SEN BIZLERIN YARDIMCISI OL,YA RABBIM! AMIN");
		
		FileOutputStream out=new FileOutputStream(excelPath);
		wb.write(out);
		

//		FileInputStream in = new FileInputStream(excelPath);
//		XSSFWorkbook workbook = new XSSFWorkbook(in);
//		XSSFSheet worksheet = workbook.getSheet("TestData");
//
//		int rowsCount = worksheet.getPhysicalNumberOfRows();
//		System.out.println(rowsCount);
//
//		XSSFCell cell = worksheet.getRow(2).getCell(2);
//		if (cell == null) {
//			cell = worksheet.getRow(2).createCell(2);
//		}
//		cell.setCellValue("Fail");
//		
//		//Fail the android line
//		cell=worksheet.getRow(1).getCell(2);
//		if(cell==null){
//			cell=worksheet.getRow(1).createCell(2);
//		}
//		cell.setCellValue("Pass");
//		//Fail the android line
//		cell=worksheet.getRow(5).getCell(2);
//		if(cell==null){
//			cell=worksheet.getRow(5).createCell(2);
//			
//		}
//		cell.setCellValue("FAIL!");
//		
//		//Wooden Spoon pass
//		cell=worksheet.getRow(4).getCell(2);
//		if(cell==null){
//			cell=worksheet.getRow(4).createCell(2);
//		}
//		cell.setCellValue("PASSED!!!");
//		
//		FileOutputStream out=new FileOutputStream(excelPath);
//		workbook.write(out);
		
		in.close();
		out.close();
	}
}
