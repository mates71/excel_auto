package udemy;

import utilities.ExcelUtils;

public class ExcelReadByUtils {

	public static void main(String[] args) {
		

		String exlPath="./src/test/resources/TestData/EmpData.xlsx";
		
		ExcelUtils.openExcelFile(exlPath, "Sheet4");
		
		int rowCount=ExcelUtils.getUsedRowsCount();
		
		System.out.println(rowCount);
		
		for( int i=0;i<rowCount;i++){
			
			String name=ExcelUtils.getCellData(i, 0);
		//	System.out.println(name);
			
			String position=ExcelUtils.getCellData(i, 1);
			
			String office=ExcelUtils.getCellData(i, 2);
			
			String ext=ExcelUtils.getCellData(i, 3);
			
			String date=ExcelUtils.getCellData(i, 4);
			
			String salary=ExcelUtils.getCellData(i, 5);
			
			System.out.println(name+" | "+position+"  | "+office+"  | "+ext+" | "+date+" | "+salary);
			
		}
		
	}

}
