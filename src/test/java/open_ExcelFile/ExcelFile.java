package open_ExcelFile;

public class ExcelFile {
/*
	String path = FileUtils.getLatestFileFromDir("Your folder path").toString();
	ExcelUtils.openExcelFile(path, "Sheet1");
	String r = ExcelUtils.getCellData(0, 0);
	System.out.println(r);
	*/
	public static void main(String[] args) {
		
		ExcelFile excelFile=new ExcelFile();
		excelFile.output(true, false);
		
	}
		public void output(boolean a, boolean b) {

		    if (a && b) {

		        System.out.println("A && B");

		    } else if (a) {

		        System.out.println("A");

		    } else {

		        if (!b) {

		            System.out.println("notB");

		        } else {

		            System.out.println("ELSE");

		        }

		    }
}}