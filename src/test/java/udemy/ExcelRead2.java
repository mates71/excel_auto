package udemy;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead2 {

	public static void main(String[] args) throws IOException {

		String exlPath = "./src/test/resources/TestData/EmpData.xlsx";

		FileInputStream input = new FileInputStream(exlPath);
		XSSFWorkbook wb = new XSSFWorkbook(input);
		XSSFSheet sh = wb.getSheet("Sheet3");
		XSSFRow row = sh.getRow(0);
		XSSFCell cell = row.getCell(0);

		int rows = sh.getPhysicalNumberOfRows();
		System.out.println(rows);

		cell = sh.getRow(1).getCell(1);
		System.out.println(cell);// Anne

		String cellData = cell.getStringCellValue();
		System.out.println(cellData);

		// print all names

		for (int i = 0; i < rows; i++) {

			String allNames = sh.getRow(i).getCell(1).toString();
			String ages = sh.getRow(i).getCell(2).toString().substring(0, 2).replace(".", "");

			int ages2 = Integer.parseInt(ages);
			System.out.println(allNames + " | " + ages2);
		}

		input.close();
	}

}
