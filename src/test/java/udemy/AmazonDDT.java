package udemy;

import java.time.LocalDateTime;
import java.util.concurrent.TimeUnit;

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

public class AmazonDDT {
	
	static WebDriver driver;
	WebElement search;
	WebElement results;
	
	String exlPath="./src/test/resources/TestData/AmazonSearchData.xlsx";

	@BeforeClass

	public static void setUp() {
		System.setProperty("webdriver.chrome.driver", 
				"/Users/musaates/Documents/Libraries/drivers/chromedriver");
		driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.get("https://www.amazon.com");
	}
	@Test
	
	public void searchTest1(){
		
		ExcelUtils.openExcelFile(exlPath, "Sheet1");
		int rowCount=ExcelUtils.getUsedRowsCount();
		
		for( int i=1;i<rowCount;i++){
			
			String execute=ExcelUtils.getCellData(i, 0);
			
			System.out.println(execute);
			
			if(execute.equalsIgnoreCase("N")){
				
				ExcelUtils.setCellData(" ", i, 2);
				ExcelUtils.setCellData("skiped", i, 2);
				continue;
			}
			
			String searchItem=ExcelUtils.getCellData(i, 1);
			
			searchFor(searchItem);
			
			String resultText=getSearchresults();
			
			int resultCount=cleanUpResultCount(resultText);
			System.out.println("resultCount :"+resultCount);
			
			
			ExcelUtils.setCellData(String.valueOf(resultCount), i, 2);
			
			if(resultCount>0){
				
				ExcelUtils.setCellData("PASS", i, 3);
			}else{
				ExcelUtils.setCellData("FAILED", i, 3);
			}
			
			String now=LocalDateTime.now().toString();
			ExcelUtils.setCellData(now, i, 4);
		}
		
		
	}
	
	public String getSearchresults(){
		
		try{
			WebDriverWait wait=new WebDriverWait(driver, 10);
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("s-result-count")));
		results=driver.findElement(By.id("s-result-count"));
		}catch(Exception noElem){
			return "0 results";
		}
		return results.getText();
		
	}
	//~~~~~~~~~~~~~~~~~~~~~~
	public int cleanUpResultCount(String resulttext){
		String[] arrResult=resulttext.split(" ");
		
		int resultCount;
		
		if(resulttext.contains(" of ")){
			resultCount=Integer.parseInt(arrResult[3].replaceAll(",", ""));
			
		}else{
			resultCount=Integer.parseInt(arrResult[0]);
		}
		
		return resultCount;
	}
	
	//~~~~~~~~~~~~~~
		public void searchFor(String item){
			search=driver.findElement(By.id("twotabsearchtextbox"));
			search.clear();
			search.sendKeys(item+Keys.ENTER);
			
		}
		
	}


