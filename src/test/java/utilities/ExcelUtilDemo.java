package utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtilDemo {

	private static XSSFWorkbook excelWorkbook;
	private static XSSFSheet excelSheet;
	private static XSSFCell cell;
	private static XSSFRow row;
	private static String excelFilePath;

	public static void openExcelFile(String path, String sheetName) {
		excelFilePath = path; // open excel file
		try {
			FileInputStream ExcelFile = new FileInputStream(path);
			excelWorkbook = new XSSFWorkbook(ExcelFile);
			excelSheet = excelWorkbook.getSheet(sheetName);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	// read test data
	public static String getCellData(int rowNum, int colNum) {
		try {
			cell = excelSheet.getRow(rowNum).getCell(colNum);
			String cellData = cell.toString();
			return cellData;
		} catch (Exception e) {
			e.printStackTrace();
			return "";
		}
	}

	// to write in to excel
	public static void setCellDAta(String value,int rowNum,int colNum) {
try{
	row=excelSheet.getRow(rowNum);
	cell=row.getCell(colNum);
	
	if(cell==null){
		cell=row.createCell(colNum);
		cell.setCellValue(value);
		
	}else{
		
		cell.setCellValue(value);
	}
FileOutputStream fileOut=new FileOutputStream(excelFilePath);
excelWorkbook.write(fileOut);
fileOut.close();
	
	}catch(Exception e){
		e.printStackTrace();;
	}
}
	
	public static int getUsedRowCount(){
		try{
			int rowCount=excelSheet.getPhysicalNumberOfRows();
			return rowCount;
			
		}catch(Exception e){
			
			e.printStackTrace();
			return 0;
		}
	
		
		
	}
}