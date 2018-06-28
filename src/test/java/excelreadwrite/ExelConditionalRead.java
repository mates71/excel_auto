package excelreadwrite;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExelConditionalRead {

	public static void main(String[] args) throws IOException {

		String excelPath = "./src/test/resources/TestData/EmpData.xlsx";
		// Open the file to read
		FileInputStream in = new FileInputStream(excelPath);
		// Let the Apache XSSFWorkbook class handle the data
		XSSFWorkbook workbook = new XSSFWorkbook(in);
		// Jump to Worksheet level
		XSSFSheet worksheet = workbook.getSheet("Sheet2");

		int rowsCount = worksheet.getPhysicalNumberOfRows();

		for (int rownum = 1; rownum < rowsCount; rownum++) {
			// read first cell value
			String execute = worksheet.getRow(rownum).getCell(0).toString();
			// if it is Y then switch to next cell and print the searchItem
			// value
			if (execute.equals("Y")) {
				String searchItem = worksheet.getRow(rownum).getCell(1).toString();
				System.out.println("Searching for " + searchItem);
			}
		}

		in.close();
	}
}
