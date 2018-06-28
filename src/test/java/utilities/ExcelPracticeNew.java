package utilities;

import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelPracticeNew {

	
	
	public static void main(String[] args) throws Exception {
		String excelFilePath ="./src/test/resources/TestData/AmazonSearchDDT.xlsx";
		/// .xlsx new version  
		
		FileInputStream in = new FileInputStream(excelFilePath); 
		//Workbook
		XSSFWorkbook wb1 = new XSSFWorkbook(in); 
		//WorkSheet
		XSSFSheet	sh1 = wb1.getSheet("Sheet1"); 
		XSSFSheet   sh2 = wb1.getSheetAt(0); 
		//Row
		XSSFRow   rw = sh1.getRow(3);
		//Cell 
		XSSFCell cell = rw.getCell(2);
		
		System.out.println(cell.toString());
		XSSFRow firstRow = sh1.getRow(0);
		
		int rowNum = sh1.getPhysicalNumberOfRows(); 
		int colNum = firstRow.getPhysicalNumberOfCells();
		
		for (int i = 0; i < rowNum; i++) {
			
			for (int j = 0; j < sh1.getRow(i).getPhysicalNumberOfCells(); j++) {
				
//				HSSFRow   rw1 = sh1.getRow(i);
//				 
//				HSSFCell cell1 = rw.getCell(j);
//				
				
				//System.out.print(cell1+"-------");
				
				System.out.print(sh1.getRow(i).getCell(j).toString() + "-------");
				
				
			}
			System.out.println("end of row "+i);
			
		}
		in.close();
	

}
	
}
