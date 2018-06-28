package udemy;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

	public static void main(String[] args) throws IOException {

		

		String path ="./src/test/resources/TestData/EmpData.xlsx";
		XSSFCell cell;

		FileInputStream input = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(input);
		XSSFSheet worksheet = workbook.getSheet("Sheet1");
//print rows
		int rows = worksheet.getPhysicalNumberOfRows();
		System.out.println("Total Rows " + rows);
//print cell
		System.out.println(worksheet.getRow(2).getCell(1));
		cell = worksheet.getRow(1).getCell(1);
		String cellData = cell.getStringCellValue();
		System.out.println(cellData);
		
		//print all names
		for(int row=1;row<rows;row++){
			String Name=worksheet.getRow(1).getCell(0).toString();			
			System.out.println(Name);
		}

	}
}
