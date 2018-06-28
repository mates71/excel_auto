package excelautomation_home;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead2 {
	public static void main(String[] args) throws IOException {

		String excelPath = "./src/test/resources/TestData/newExcel.xlsx";

		FileInputStream fis = new FileInputStream(excelPath);

		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sh = wb.getSheet("Sheet1");
		XSSFRow row = sh.getRow(0);
		XSSFCell cell = row.getCell(0);

		int countRow = sh.getPhysicalNumberOfRows();
		int countCell = row.getPhysicalNumberOfCells();

		System.out.println("total row :" + countRow);
		System.out.println("total cell :" + countCell);

		System.out.println("I love so much my " + sh.getRow(6).getCell(1));
		System.out.println("I am reading the " + sh.getRow(1).getCell(1));

		for (int i = 0; i < countRow; i++) {

			String execute = sh.getRow(i).getCell(0).toString();
			String searhItem = sh.getRow(i).getCell(1).toString();

			System.out.println(execute + "-->" + searhItem);

		}

	}

}
