package utilities;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class ExcelPracticeUltimate {

	public static void main(String[] args) throws Exception {
		
		String excelFilePath = "./src/test/resources/TestData/AmazonSearchData.xls";

		/// print out 3rd row 5th column 
		/// print out each and everything using loop if the flag Y
		
	
		FileInputStream fis = new FileInputStream(excelFilePath); 
		//Workbook 
		Workbook wb = WorkbookFactory.create(fis); 
		//WorkSheet
		Sheet sh = wb.getSheetAt(0);
		//Row
		Row row = sh.getRow(0); 
		//Cell 
		Cell cell = row.getCell(1);
		
		System.out.println(" print r3,c5 "+sh.getRow(3).getCell(1));
		
		for (int i = 0; i < sh.getPhysicalNumberOfRows(); i++) {
			
			if(sh.getRow(i).getCell(0).toString().equalsIgnoreCase("Y")){
				
				for (int j = 0; j < sh.getRow(i).getPhysicalNumberOfCells(); j++) {
					System.out.print(sh.getRow(i).getCell(j).toString()+"==========");
					
				}
				System.out.println();
			}else{
				System.out.println("row number "+i+" skiped");
			}
			
		}
		
		/*for (int i = 1; i < sh.getPhysicalNumberOfRows(); i++) {
			
			// if the excecution flag is Y 
			
			if(sh.getRow(i).getCell(0).toString().equalsIgnoreCase("Y")) {
			
				for (int j = 0; j < sh.getRow(i).getPhysicalNumberOfCells(); j++) {
					System.out.print(sh.getRow(i).getCell(j).toString()+"------");
				}
				System.out.println();
				
			}else{
				System.out.println("row number "+ i + " is skipped");
				
			}
			
			
		}*/
	}
		

	}



