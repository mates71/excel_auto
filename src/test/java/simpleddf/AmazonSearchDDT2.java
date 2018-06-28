package simpleddf;

import java.io.FileInputStream;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.BeforeClass;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class AmazonSearchDDT2 {
	static WebDriver driver;

	@BeforeClass
	public static void setUp() {
		System.setProperty("webdriver.chrome.driver", "/Users/musaates/Documents/Libraries/drivers/chromedriver");

		driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.get("https://www.amazon.com");
	}

	@Test
	public void test() throws Exception{
		//Open Excel File
				String excelPath="./src/test/resources/TestData/AmazonSearchData.xlsx";
				Thread.sleep(2000);
				FileInputStream in=new FileInputStream(excelPath);
				//Let the Apache XSSFWorkbook class handle the data
				XSSFWorkbook workbook=new XSSFWorkbook(in);
				//Jump to Worksheet level
				XSSFSheet worksheet=workbook.getSheet("Sheet1");
				//get rows count
				//Start your for loop , make sure it runs for all rows
				//if execute row says Y then
					//Store the search item into a variable and sendkeys to search inputfield {
				
					driver.findElement(By.id("twotabsearchtextbox")).sendKeys("Cucumber book"+Keys.ENTER);
					String results=driver.findElement(By.id("s-result-count")).getText();
					
					System.out.println(results);
					int numberOfResults=cleanupSearchResultsCount(results);
					
					if(numberOfResults > 0){
						System.out.println("PASS");
						//Write to excel status cell
					}else{
						System.out.println("FAIL");
						//Write to excel status cell
					}
			  //} CLOSE for LOOP
				//Save excel file changes, close in,out,workbook.
					//quit the driver
			}

			public int cleanupSearchResultsCount(String results){
				
				String[] resultArr=results.split(" ");
				int resultsCount= Integer.parseInt(resultArr[2].replace(",", ""));
				System.out.println(resultsCount);
				return resultsCount;
				
				
			}
			
}
