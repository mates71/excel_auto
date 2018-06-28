package excelreadwrite;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {

	public static void main(String[] args) throws Exception {

		String excelPath = "./src/test/resources/TestData/EmpData.xlsx";
		// Open the file to read
		FileInputStream in = new FileInputStream(excelPath);
		// Let the Apache XSSFWorkbook class handle the data
		XSSFWorkbook workbook = new XSSFWorkbook(in);
		// Jump to Worksheet level
		XSSFSheet worksheet = workbook.getSheet("Sheet2");

		int rowsCount = worksheet.getPhysicalNumberOfRows();

		System.out.println("rowsCount:" + rowsCount);

		XSSFCell cell = worksheet.getRow(1).getCell(2);
		if (cell == null) {
			cell = worksheet.getRow(1).createCell(2);
		}
		cell.setCellValue("Fail");

		// Fail the android line
		cell = worksheet.getRow(5).getCell(2);
		if (cell == null) {
			cell = worksheet.getRow(5).createCell(2);
		}
		cell.setCellValue("Pass");
		// SAve changes
		FileOutputStream out = new FileOutputStream(excelPath);
		workbook.write(out);

		in.close();
		out.close();
	}
}
