package udemy;

import java.io.IOException;

import utilities.ExcelUtils;

public class WriteToExcel2 {

	public static void main(String[] args) throws IOException {

		String excelPath = "./src/test/resources/TestData/Employees.xlsx";
		
		ExcelUtils.openExcelFile(excelPath, "Sheet1");

		int rowCount = ExcelUtils.getUsedRowsCount();
		System.out.println(rowCount);

		for (int i = 0; i < rowCount; i++) {

			String name = ExcelUtils.getCellData(i, 0);
			String position = ExcelUtils.getCellData(i, 1);
			String office = ExcelUtils.getCellData(i, 2);
			String ext = ExcelUtils.getCellData(i, 3);
			String Date = ExcelUtils.getCellData(i, 4);
			String Salary = ExcelUtils.getCellData(i, 5);

			System.out.println(name + " | " + position + " | " + office + " | " + ext + " | " + Date + " | " + Salary);

			

		}

	}

}

/*
 * FileInputStream fis = new FileInputStream(excelPath); XSSFWorkbook wb = new
 * XSSFWorkbook(fis); XSSFSheet wsh = wb.getSheet("Sheet1");
 * 
 * XSSFRow row = wsh.getRow(0); XSSFCell cell = row.getCell(0);
 * 
 * String cellValue = wsh.getRow(0).getCell(2).toString();
 * //System.out.println(cellValue);
 * 
 * int rowsCount = wsh.getPhysicalNumberOfRows();
 * //System.out.println(rowsCount);
 * 
 * for (int i = 0; i < rowsCount; i++) {
 * 
 * String firstName = wsh.getRow(i).getCell(0).toString(); String lastName =
 * wsh.getRow(i).getCell(1).toString(); String salary =
 * wsh.getRow(i).getCell(2).toString(); String hireDate =
 * wsh.getRow(i).getCell(3).toString();
 * 
 * System.out.println(firstName + " | " + lastName + " | " + salary + " | " +
 * hireDate); }
 * 
 * }
 */
