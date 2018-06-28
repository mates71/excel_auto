package excelautomation_home;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite2 {

	public static void main(String[] args) throws IOException {
		
		String excelPath="./src/test/resources/TestData/newExcel.xlsx";
		
		FileInputStream fis=new FileInputStream(excelPath);
		
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sh=wb.getSheet("Sheet1");
		XSSFRow row=sh.getRow(0);
//		XSSFCell cell=row.getCell(0);
		
		int countRow=sh.getPhysicalNumberOfRows();
		int cellCount=row.getPhysicalNumberOfCells();
		
		for (int i = 1; i < countRow; i++) {
			
		
		
		XSSFCell cell=sh.getRow(i).getCell(2);
		if(cell==null){
			cell=sh.getRow(i).createCell(2);
			
		}
		cell.setCellValue("PASS");
		
		FileOutputStream fos=new FileOutputStream(excelPath);
		wb.write(fos);
		
		fis.close();
		fos.close();
		
		}
		
		
		
	}
}
