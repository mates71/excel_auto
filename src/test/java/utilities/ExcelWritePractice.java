package utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelWritePractice {

	
public static void main(String[] args) throws Exception {
		
		/// .xls older version  
			String excelFilePath ="./src/test/resources/TestData/AmazonSearchData.xls";
			
			FileInputStream in=new FileInputStream(excelFilePath);
		
			HSSFWorkbook wb=new HSSFWorkbook(in);
			HSSFSheet sh=wb.getSheet("Sheet1");
			
			HSSFRow row=sh.getRow(3);
			HSSFCell cell=row.getCell(4);
			
			int rowCount=sh.getPhysicalNumberOfRows();
			int cellCount =row.getPhysicalNumberOfCells();
			
			
			if(cell==null){
				row.createCell(3);
				cell.setCellValue("Musa");
			}else{
				
				cell.setCellValue("Musa");
			}
		
			FileOutputStream out=new FileOutputStream(excelFilePath);
			wb.write(out);
			out.close();
}
//			FileInputStream in = new FileInputStream(excelFilePath);
//			
//			//Workbook
//			HSSFWorkbook wb1 = new HSSFWorkbook(in); 
//			//WorkSheet
//			HSSFSheet	sh1 = wb1.getSheet("Sheet1"); 
//			//HSSFSheet   sh2 = wb1.getSheetAt(0); 
//			//Row
//			
//			HSSFRow   rw = sh1.getRow(1);
//			HSSFCell  cell = rw.getCell(1);
//			if(cell==null){
//				rw.createCell(1);
//				cell.setCellValue("HELLO");
//			}else{
//				cell.setCellValue("HELLO");
//			}
//			FileOutputStream out = new FileOutputStream(excelFilePath);
//			wb1.write(out);
//			
//			out.close();
//			
			//Cell 
			//HSSFCell cell = rw.getCell(2);

			
			}	
	


