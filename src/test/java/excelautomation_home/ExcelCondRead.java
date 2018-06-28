package excelautomation_home;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCondRead {

	public static void main(String[] args) throws IOException {
		
String excelPath="./src/test/resources/TestData/newExcel.xlsx";
		
		FileInputStream fis=new FileInputStream(excelPath);
		
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sh=wb.getSheet("Sheet1");
		XSSFRow row=sh.getRow(0);
		XSSFCell cell=row.getCell(0);

		int rowCount=sh.getPhysicalNumberOfRows();
		
		for (int i = 0; i < rowCount; i++) {
			
			String Execute=sh.getRow(i).getCell(0).toString();
			if(Execute.equals("N")){
				String searcItem=sh.getRow(i).getCell(1).toString();
				System.out.println(searcItem);
			}
			
		}
		fis.close();
	}

}
