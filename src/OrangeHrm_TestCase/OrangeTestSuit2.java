package OrangeHrm_TestCase;

import org.testng.annotations.Test;

import AppUtils.UtilsDemo;
import AppUtils.XL_Utils_Next;
import OrangeHRM_Libreary.Login;

public class OrangeTestSuit2 extends UtilsDemo
{

	String tcfile = "E:\\selinium\\myproject\\OrangeHrm_dHybride_FrameWork\\DataFile\\Hybride_Testing.xlsx";
	String tcsheet ="TestCase2";
	String tssheet = "TestStep2";
	
	String tcid,tcexeflag;
	String tstcid,Keyword	;
	
	Login sanju = new Login();
	String adminuid,adminpwd;
	
	String stepres;
	boolean res =false;
	
	@Test
	public void checkOrangeHrm() throws Throwable 
	{
		
     int tccount =	XL_Utils_Next.getRowcount(tcfile, tcsheet);
     int tscount =  XL_Utils_Next.getRowcount(tcfile, tssheet);
     
     
     
     
     for(int i=1;i<=tccount;i++)
     {
    	tcid= XL_Utils_Next.getStringData(tcfile, tcsheet, i, 0);    	
    	tcexeflag = XL_Utils_Next.getStringData(tcfile, tcsheet, i, 2); 
    	
    	if(tcexeflag.equalsIgnoreCase("y"))
    	{
    		
    		for(int j=1;j<=tscount;j++)
    		{
    			
    		 tstcid =	XL_Utils_Next.getStringData(tcfile, tssheet, j, 0);
    		 if(tstcid.equalsIgnoreCase(tcid))
    		 {
    			 Keyword = XL_Utils_Next.getStringData(tcfile, tssheet, j, 4);
    			 switch (Keyword.toLowerCase())
    		    {
    			 case "adminlogin":
			   adminuid = XL_Utils_Next.getStringData(tcfile, tssheet, j, 5);
			   adminpwd = XL_Utils_Next.getStringData(tcfile, tssheet, j, 6);
			    sanju.login(adminuid, adminpwd);
    		     res =	sanju.isdisplayed();
    		     break;
    			 case "logout":
    		    res = sanju.logout();
			    break;
				}
    			
    			 //code to update step result
    			 
    			 if(res)
    			 {
    				 
    				 stepres = "pass";
    				 XL_Utils_Next.getsetData(tcfile, tssheet, j, 3, stepres);    				 
    				 XL_Utils_Next.FillGreenColour(tcfile, tssheet, j, 3);
    			 }else
    			 {   
    				 stepres = "fail";
    				 XL_Utils_Next.getsetData(tcfile, tssheet, j, 3, stepres);
    				 
    			 }
    				 
    			 
    		 }
    			
    			
    			
    		}
    		
    		
    	}else
    	{
    		XL_Utils_Next.getsetData(tcfile, tcsheet, i, 3, "Block");
    		XL_Utils_Next.FillRedColour(tcfile, tcsheet, i, 3);
    	}
    		
     }
     
     
		
	}
	
	
	
}
