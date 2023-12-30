package AppUtils;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.BeforeTest;

public class UtilsDemo 
{
	public static WebDriver driver;
	public static String url = "http://orangehrm.qedgetech.com";
			
			
	@BeforeSuite
	public static void LanchApp() 
	{
		System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
		
		driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		driver.manage().deleteAllCookies();
		driver.manage().window().maximize();
		driver.get(url);
		
		
	}
	
    @AfterSuite
	public  static void CloseApp() 
	{
		
		driver.close();
		
	}
	
	
	
	
	}

		
	
	
	
	
	
	
	
	
	
	
	

