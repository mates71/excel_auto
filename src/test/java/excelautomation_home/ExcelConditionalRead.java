package excelautomation_home;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelConditionalRead {

	public static void main(String[] args) throws IOException {
		
		String excelPath="/Users/musaates/Desktop/EmpData.xlsx";
		FileInputStream in=new FileInputStream(excelPath);
		XSSFWorkbook workbook=new XSSFWorkbook(in);
		XSSFSheet worksheet=workbook.getSheet("TestData");
		
		int rowsCount=worksheet.getPhysicalNumberOfRows();
	
		
		for(int rownum=1;rownum<rowsCount;rownum++){
			//read first cell value
			String execute=worksheet.getRow(rownum).getCell(0).toString();
			//if it is Y then switch to next cell and print the searchItem value
			if(execute.equals("Y")){
				String searchItem=worksheet.getRow(rownum).getCell(1).toString();
				System.out.println("searching for : "+searchItem);
			}
		}
in.close();
	}
}

