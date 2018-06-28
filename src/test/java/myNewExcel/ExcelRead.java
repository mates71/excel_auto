package myNewExcel;

import java.util.List;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class ExcelRead {

	static WebDriver driver;

	@BeforeClass
	public static void setUp() {
		System.setProperty("webdriver.chrome.driver", "/Users/musaates/Documents/Libraries/drivers/chromedriver");
		driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.get("https://learn.cybertekschool.com/courses/2/users");

	}

	@Test(priority = 0)
	public void loginTest() throws InterruptedException {
		driver.findElement(By.id("pseudonym_session_unique_id")).sendKeys("musaates");
		driver.findElement(By.id("pseudonym_session_password")).sendKeys("#Tmsa2008");
		Thread.sleep(2000);
		driver.findElement(By.xpath("(//button[@type='submit'])[1]")).click();

	}

	@Test(priority = 1)
	public void readExcel() {
		WebElement table = driver.findElement(By.xpath("//tbody[@class='collectionViewItems']"));
		List<WebElement> rows = table.findElements(By.tagName("tr"));

		System.out.println(rows.size());

		for (int i = 0; i < rows.size(); i++) {

			System.out.println(rows.get(i).getText());
		}
		System.out.println();
	}
	
	@Test(priority = 2)
	public void readExcel2(){
		
		WebElement table = driver.findElement(By.xpath("//tbody[@class='collectionViewItems']"));
		List<WebElement> rows = table.findElements(By.tagName("tr"));

		for(WebElement row :rows){
			
			System.out.println(row.getText());
		}
		
	}

}
