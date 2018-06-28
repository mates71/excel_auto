package myNewExcel;

import java.util.List;
import java.util.concurrent.TimeUnit;

import org.junit.AfterClass;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import utilities.ExcelUtils;

public class ExcelNew {

	static WebDriver driver;
	static String excelPath = "./src/test/resources/TestData/Employees.xlsx";
	static String url="https://letskodeit.teachable.com/p/practice";
	@BeforeClass
	public static void setUp() {
		System.setProperty("webdriver.chrome.driver", "/Users/musaates/Documents/Libraries/drivers/chromedriver");
		driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.get("https://www.w3schools.com/html/html_tables.asp");
		
		String xlPath="/src/test/resources/TestData/AmazonSearchData.xlsx";
		

	}

	@Test
	public void tableRead() {

		WebElement table = driver.findElement(By.id("customers"));

		List<WebElement> rows = table.findElements(By.tagName("tr"));

		for (int i = 0; i < rows.size(); i++) {
		
			System.out.println(rows.get(i).getText() + " | ");
		}
		System.out.println();
		
	}
	
	@Test
	public void writeToExl(){
		
		int rowCount=ExcelUtils.getUsedRowsCount();
		System.out.println(rowCount);
		
		ExcelUtils.openExcelFile(excelPath, "Sheet4");
		
		for(int i=0;i<rowCount;i++){
		
			String d1=ExcelUtils.getCellData(i, 0);
			System.out.println(d1);
			
		ExcelUtils.setCellData(d1, i, 0);
		ExcelUtils.setCellData("Contact", i, 1);
		ExcelUtils.setCellData("Country", i, 2);
		
		
		}	
	}
	
	@Test(enabled=false)
	public void tableRead2() {
		driver.get(url);
		WebElement table2=driver.findElement(By.id("product"));
		List<WebElement> rows=table2.findElements(By.tagName("tr"));
		
		for (WebElement webElement : rows) {
			System.out.println(webElement.getText());
		
		}
	}
	@AfterClass
	public void tearDown(){
		//driver.quit();
		
	}
}
