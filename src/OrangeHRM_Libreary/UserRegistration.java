package OrangeHRM_Libreary;

import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;

public class UserRegistration extends AdminUtils
{

	public boolean addUser(String role ,String empname , String uname, String pword) throws InterruptedException 
	{
		
		driver.findElement(By.linkText("Admin")).click();
		driver.findElement(By.linkText("User Management")).click();
		driver.findElement(By.linkText("Users")).click();
		driver.findElement(By.id("btnAdd")).click();
		
		Select sanju = new Select(driver.findElement(By.id("systemUser_userType")));
		sanju.selectByVisibleText(role);
		
		driver.findElement(By.id("systemUser_employeeName_empName")).sendKeys(empname);
		driver.findElement(By.id("systemUser_userName")).sendKeys(uname);
		driver.findElement(By.id("systemUser_password")).sendKeys(pword);
		driver.findElement(By.id("systemUser_confirmPassword")).sendKeys(pword);
		
		Thread.sleep(2000);
		driver.findElement(By.id("btnSave")).click();
		
		driver.findElement(By.id("searchSystemUser_userName")).sendKeys(uname);
		driver.findElement(By.id("searchBtn")).click();
		
		
	WebElement usertable =	driver.findElement(By.id("resultTable"));
		
	        List<WebElement>rows = usertable.findElements(By.tagName("tr"));
	         
	        boolean isUserdisplayed = false;
	        
	        for(int i=1;i<rows.size();i++)
	        {
	        List<WebElement>cols =	rows.get(i).findElements(By.tagName("td"));
	        
	            if(cols.get(1).getText().equals(uname))
	            {
	            	
	            	return true;
	            }
	        
	        
	        
	        
	        }
	
	
	      return isUserdisplayed;
	
		
	}
	
	
	
}
