package excelautomation_home;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import utilities.ExcelUtils;

public class AmazonSearchResultDDT {
	static WebDriver driver;

	@BeforeClass
	public static void setUp() {
		System.setProperty("webdriver.chrome.driver", "/Users/musaates/Documents/Libraries/drivers/chromedriver");
		driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.get("https://www.amazon.com");
	}

	@Test
	public void test() throws IOException {
		String excelPath ="./src/test/resources/TestData/newAmazon.xlsx";
						
		FileInputStream in = new FileInputStream(excelPath);
		XSSFWorkbook workbook = new XSSFWorkbook(in);
		XSSFSheet worksheet = workbook.getSheet("Sheet1");

		driver.findElement(By.id("twotabsearchtextbox")).sendKeys("Cucumber book" + Keys.ENTER);
		String results = driver.findElement(By.id("s-result-count")).getText();

		System.out.println(results);

		int numberOfResut = cleanupSearchResult(results);
		System.out.println(numberOfResut);

		if (numberOfResut > 0) {
			ExcelUtils.setCellData("PASS", 2, 2);
		
		} else {
			
			if (numberOfResut > 0) {
				
				ExcelUtils.setCellData("FAIL", 2, 2);
				
				
			}

			FileOutputStream fos = new FileOutputStream(excelPath);
			workbook.write(fos);

			in.close();
			fos.close();
		}

	}

	public int cleanupSearchResult(String results) {

		String[] resultArr = results.split(" ");

		int resultCount = Integer.parseInt(resultArr[2].replace(",", ""));

		return resultCount;

	}
}