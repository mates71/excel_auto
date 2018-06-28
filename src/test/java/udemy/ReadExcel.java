package udemy;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) throws IOException {

		String excelPath = "./src/test/resources/TestData/newExcel.xlsx";

		FileInputStream fis = new FileInputStream(excelPath);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet wsh = wb.getSheet("Sheet2");

		XSSFRow row = wsh.getRow(0);
		XSSFCell cell = row.getCell(0);

		// String cellValue = wsh.getRow(0).getCell(2).toString();
		// System.out.println(cellValue);

		int rowCount = wsh.getPhysicalNumberOfRows();
		System.out.println(rowCount);

		for (int i = 1; i < rowCount; i++) {

			String SALARY = wsh.getRow(i).getCell(2).toString();
			// System.out.println(SALARY);
			int SALARY2 = Integer.parseInt(wsh.getRow(i).getCell(2).toString().replace(".", ""));// to convert to int

			// System.out.println(SALARY2);

			if (SALARY2 <= 25000) {
				XSSFCell STATUS = wsh.getRow(i).getCell(4);
				if (STATUS == null) {
					STATUS = wsh.getRow(i).createCell(4);

				}
				STATUS.setCellValue("Not GOOD");

			}
			if (SALARY2 > 25000) {
				XSSFCell STATUS = wsh.getRow(i).getCell(4);
				if (STATUS == null) {
					STATUS = wsh.getRow(i).createCell(4);

				}
				STATUS.setCellValue("GOOD");

				FileOutputStream fos = new FileOutputStream(excelPath);
				wb.write(fos);

				fis.close();
				fos.close();
			}

		}
	}
}
