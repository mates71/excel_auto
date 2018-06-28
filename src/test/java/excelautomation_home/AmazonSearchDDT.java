package excelautomation_home;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.BeforeClass;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class AmazonSearchDDT {

	static WebDriver driver;
	
	@BeforeClass
	public static void setUp(){
		System.setProperty("webdriver.chrome.driver",
				"/Users/musaates/Documents/Libraries/drivers/chromedriver");
		driver=new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.get("https://www.amazon.com");
	}
@Test
public void test() throws IOException{
	String excelPath="./src/test/resources/TestData/AmazonSearchData.xlsx";
	FileInputStream in=new FileInputStream(excelPath);
	XSSFWorkbook workbook=new XSSFWorkbook(in);
	XSSFSheet worksheet=workbook.getSheet("Sheet1");
	
	driver.findElement(By.id("twotabsearchtextbox")).sendKeys("Cucumber book"+Keys.ENTER);
	String results=driver.findElement(By.id("s-result-count")).getText();
	System.out.println(results);
	
	int numberOfResults=cleanupSearchResultsCount(results);
	if(numberOfResults>0){
		System.out.println("Pass");
	}else{
		System.out.println("Fail");
	}
	
}
private int cleanupSearchResultsCount(String results) {
	
	String[] resultArr=results.split(" ");
	System.out.println(resultArr[1]);
	int resultCount=Integer.parseInt(resultArr[2].replace(",", ""));

	
	return resultCount;
}


}
