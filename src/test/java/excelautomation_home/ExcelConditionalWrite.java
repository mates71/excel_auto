package excelautomation_home;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelConditionalWrite {

	public static void main(String[] args) throws IOException {
		
		String excelPath= "/Users/musaates/Desktop/EmpData.xlsx";
		
		FileInputStream in=new FileInputStream(excelPath);
		XSSFWorkbook workbook=new XSSFWorkbook(in);
		XSSFSheet worksheet=workbook.getSheet("TestData");
		
		int rowsCount=worksheet.getPhysicalNumberOfRows();
		
		for(int rownum=1;rownum<rowsCount;rownum++){
			String searchItem=worksheet.getRow(rownum).getCell(1).toString();
			
			if(searchItem.contains("Cucumber")){
				XSSFCell status=worksheet.getRow(rownum).getCell(2);
				if(status==null){
					status=worksheet.getRow(rownum).createCell(2);
					status.setCellValue("Pass");
					break;
				}
			}
			//Save changes
			FileOutputStream ExcelWriteFile=new FileOutputStream(excelPath);
			workbook.write(ExcelWriteFile);
			
			ExcelWriteFile.close();
			workbook.close();
			in.close();
		}
		
		

	}

}
