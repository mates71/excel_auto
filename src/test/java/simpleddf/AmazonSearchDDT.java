package simpleddf;

import java.time.LocalDateTime;
import java.util.concurrent.TimeUnit;

import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import utilities.ExcelUtils;

public class AmazonSearchDDT {
	static WebDriver driver;
	WebElement search;
	WebElement results;

	@BeforeClass

	public static void setUp() {
		System.setProperty("webdriver.chrome.driver", 
				"/Users/musaates/Documents/Libraries/drivers/chromedriver");
		driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.get("https://www.amazon.com");
	}

	String excelFilePath ="./src/test/resources/TestData/AmazonSearchDDT.xlsx";
	  

	@Test
	public void searchTest() throws Exception {
		
	ExcelUtils.openExcelFile(excelFilePath, "Sheet2");
	int rowsCount = ExcelUtils.getUsedRowsCount();
	for (int rownum =  1; rownum <  rowsCount; rownum++){
		
	String execute=ExcelUtils.getCellData(rownum,0);
	if(execute.equalsIgnoreCase("Y")){
		
	ExcelUtils.setCellData(" ", rownum, 2);
	ExcelUtils.setCellData("Skipped", rownum, 3);
	continue;
	}
	String searchItem = ExcelUtils.getCellData(rownum, 1);
	
	searchFor(searchItem);
	String resultText = getSearchResults();
	int resultCount = cleanUpResultsCount(resultText);
	System.out.println(resultCount);
	
	
	//   write result count to excel
	ExcelUtils.setCellData(String.valueOf(resultCount), rownum, 2);
	//   report status into excel
	if   (resultCount >  0)   
	{
	System.out.println("Pass");
	ExcelUtils.setCellData("Pass", rownum, 3);
	} else {
		System.out.println("Fail");
		ExcelUtils.setCellData("Fail", rownum, 3);
	}
	//   write current time to excel
	String now =  LocalDateTime.now().toString();
	ExcelUtils.setCellData(now, rownum, 4);
	}

	}


	/** * @throws InterruptedException * 
	 * @throws Exception */
	public String getSearchResults() throws Exception {
		
		
		Thread.sleep(2222);
		try {
			WebDriverWait wait = new WebDriverWait(driver, 10);
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("s-result-count")));
			results = driver.findElement(By.id("s-result-count"));
		} catch (Exception noElem) {
			return "0 results";
		}
		return results.getText();
	}

	/** * */
	public int cleanUpResultsCount(String resultText) {
		String[] arrResult = resultText.split(" ");
		int resultsCount;
		if (resultText.contains(" of ")) {
			resultsCount = Integer.parseInt(arrResult[3].replace(",", ""));
		} else {
			resultsCount = Integer.parseInt(arrResult[0]);
		}
		return resultsCount;
	}

	/** * */
	public void searchFor(String item) {
		
		search = driver.findElement(By.id("twotabsearchtextbox"));
		search.clear();
		search.sendKeys(item + Keys.ENTER);
	}
	@AfterClass
	public static void cleanUp() {
		driver.quit();
	}
}