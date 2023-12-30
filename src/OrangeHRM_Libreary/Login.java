package OrangeHRM_Libreary;

import org.openqa.selenium.By;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;

import AppUtils.UtilsDemo;

public class Login extends UtilsDemo
{

	public void login(String ui, String pwd)
	{
		//driver.findElement(By.id("txtUsername")).clear();
		driver.findElement(By.id("txtUsername")).sendKeys(ui);
		
	//	driver.findElement(By.id("txtPassword")).clear();
		driver.findElement(By.id("txtPassword")).sendKeys(pwd);
		driver.findElement(By.id("btnLogin")).click();
		
		
	}
	
	

	public boolean logout() throws Throwable 
	{
		driver.findElement(By.partialLinkText("Welcome")).click();
		Thread.sleep(2000);
		driver.findElement(By.linkText("Logout")).click();
		if(driver.findElement(By.id("btnLogin")).isDisplayed())
		{
			return true;
		}else
		{
			return false;
		}
					
		
	}
			
		
	
	
	 public boolean isdisplayed()
	 {
		if (driver.findElement(By.linkText("Admin")).isDisplayed())
		{
		 return true; 
		}
		{
			return false;
		}
	}
	
	 
	 
	 
	 
	 public boolean isErrMsgDisplayed()
	 {
		String erromsg;
		erromsg = driver.findElement(By.id("spanMessage")).getText();
		if(erromsg.toLowerCase().contains("invalid") || erromsg.toLowerCase().contains("empty"))
		{
			return true;
		}else
		{
		 return false;
		 
	 }
	 
	 
	 }
	 
	
}
