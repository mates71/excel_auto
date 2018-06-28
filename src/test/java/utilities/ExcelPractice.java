package utilities;

import java.io.FileInputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelPractice {

	public static void main(String[] args) throws Exception {
		
		String excelFilePath ="./src/test/resources/TestData/AmazonSearchDDT.xls";

		/// .xls older version  
	
		FileInputStream in = new FileInputStream(excelFilePath); 
		//Workbook
		HSSFWorkbook wb1 = new HSSFWorkbook(in); 
		//WorkSheet
		HSSFSheet	sh1 = wb1.getSheet("Sheet1"); 
		HSSFSheet   sh2 = wb1.getSheetAt(0); 
		//Row
		HSSFRow   rw = sh1.getRow(3);
		//Cell 
		HSSFCell cell = rw.getCell(2);
		
		System.out.println(cell.toString());
		HSSFRow firstRow = sh1.getRow(0);
		
		int rowNum = sh1.getPhysicalNumberOfRows(); 
		int colNum = firstRow.getPhysicalNumberOfCells();
		
		for (int i = 0; i < rowNum; i++) {
			
			for (int j = 0; j < sh1.getRow(i).getPhysicalNumberOfCells(); j++) {
				
//				HSSFRow   rw1 = sh1.getRow(i);
//				 
//				HSSFCell cell1 = rw.getCell(j);
//				
				
				//System.out.print(cell1+"-------");
				
				System.out.print(sh1.getRow(i).getCell(j).toString() + "=======");
				
				
			}
			System.out.println("end of row "+i);
			
		}
		
		in.close();
		
		
		/// .xlsx
		//Workbook
		//WorkSheet
		//Row
		//Cell 
			
		
		

	}

}
