package OrangeHRM_Libreary;

import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import AppUtils.UtilsDemo;

public class EmpRegistration extends UtilsDemo
{

public boolean addemployee(String fname, String lname)
	
	{
		driver.findElement(By.linkText("PIM")).click();
		driver.findElement(By.linkText("Add Employee")).click();
		
		driver.findElement(By.id("firstName")).sendKeys(fname);
		driver.findElement(By.id("lastName")).sendKeys(lname);
		String empid;	
	   empid = driver.findElement(By.id("employeeId")).getAttribute("value");
	   
	   driver.findElement(By.id("btnSave")).click();
	   driver.findElement(By.linkText("Employee List")).click();
	   
	   driver.findElement(By.id("empsearch_id")).sendKeys(empid);
	   driver.findElement(By.id("searchBtn")).click();
	 
	   WebElement ram;
	  ram =  driver.findElement(By.id("resultTable"));
	   
	  List<WebElement>rows = ram.findElements(By.tagName("tr"));
	    
	  boolean isempdisplayed = false;
	  
	  for(int i=1;i<rows.size();i++)
	  {
	List<WebElement>cols =	rows.get(i).findElements(By.tagName("td"));
	       if( cols.get(1).getText().equals(empid));
	       {
	    	   
	    	   isempdisplayed = true;
	    	   
	       }
	  } 
           
	   return isempdisplayed;    
	
	  
	   
	}
		
	
	
	
	
	
	
	
	
}
