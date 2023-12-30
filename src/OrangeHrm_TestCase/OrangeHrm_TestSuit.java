package OrangeHrm_TestCase;

import org.testng.annotations.Test;

import AppUtils.UtilsDemo;
import AppUtils.XL_Utils_Next;
import OrangeHRM_Libreary.EmpRegistration;
import OrangeHRM_Libreary.Login;

public class OrangeHrm_TestSuit extends UtilsDemo {

	String tcfile = "E:\\selinium\\myproject\\OrangeHrm_dHybride_FrameWork\\DataFile\\Hybride_Testing.xlsx";
	String tcsheet = "TestCases";
	String tssheet = "TestSteps";
 
	@Test
	public void checkOrangeHRM() throws Throwable {

		int tccount = XL_Utils_Next.getRowcount(tcfile, tcsheet);
		int tscount = XL_Utils_Next.getRowcount(tcfile, tssheet);

		String tcid, tcexeflag;
		String tstcid, Keyword;

		
		Login sanju = new Login();
		EmpRegistration ram = new EmpRegistration();
		
		String adminuid, adminpwd;
		String admin1uid,admin1pwd;
		
		String fname, lname;

		String stepres,tcres;

		boolean res = false;

		for (int i = 1; i <= tccount; i++) {

			tcid = XL_Utils_Next.getStringData(tcfile, tcsheet, i, 0);
			tcexeflag = XL_Utils_Next.getStringData(tcfile, tcsheet, i, 2);

			if (tcexeflag.equalsIgnoreCase("y")) 
			{

				for (int j = 1; j <= tscount; j++) {

					tstcid = XL_Utils_Next.getStringData(tcfile, tssheet, j, 0);
					if (tstcid.equalsIgnoreCase(tcid)) {
						Keyword = XL_Utils_Next.getStringData(tcfile, tssheet, j, 4);
						
						switch (Keyword.toLowerCase()) 
						{
						case "adminlogin":
							adminuid = XL_Utils_Next.getStringData(tcfile, tssheet, j, 5);
							adminpwd = XL_Utils_Next.getStringData(tcfile, tssheet, j, 6);
							sanju.login(adminuid, adminpwd);
							res = sanju.isdisplayed();
							break;
							
						case "logout":
							res = sanju.logout();
							break;
							
						case "invalidlogin":
							
						//	System.out.println("rtt");
							admin1uid = XL_Utils_Next.getStringData(tcfile, tssheet, j, 5);
							admin1pwd = XL_Utils_Next.getStringData(tcfile, tssheet, j, 6);
							sanju.login(admin1uid, admin1pwd);
							res = sanju.isErrMsgDisplayed();
							break;
						case "empreg":
				       		fname =	XL_Utils_Next.getStringData(tcfile, tssheet, j, 5);
				       		lname = XL_Utils_Next.getStringData(tcfile, tssheet, j, 6);
						res =	ram.addemployee(fname, lname);
							break;
							
						}

						// code to update step result

						if (res) {

							stepres = "pass";
							XL_Utils_Next.getsetData(tcfile, tssheet, j, 3, stepres);
							XL_Utils_Next.FillGreenColour(tcfile, tssheet, j, 3);
						} else {
							stepres = "fail";
							XL_Utils_Next.getsetData(tcfile, tssheet, j, 3, stepres);
							XL_Utils_Next.FillRedColour(tcfile, tssheet, j, 3);
						}

						
						//code to update Testcase Result
						tcres= XL_Utils_Next.getStringData(tcfile, tcsheet, i, 3);
						if(!tcres.equalsIgnoreCase("fail"))
						{
							XL_Utils_Next.getsetData(tcfile, tcsheet, i, 3, stepres);
						}
						
						
						tcres = XL_Utils_Next.getStringData(tcfile, tssheet, i, 3);
						if(tcres.equalsIgnoreCase("pass"));
						{
							XL_Utils_Next.FillGreenColour(tcfile, tcsheet, i, 3);
						}
						{
							XL_Utils_Next.FillRedColour(tcfile, tssheet, i, 3);
						}
						
						
						
					}

				}

			}else {

						XL_Utils_Next.getsetData(tcfile, tcsheet, i, 3, "Blocked");
						XL_Utils_Next.FillRedColour(tcfile, tcsheet, i, 3);
			}
			}

		}

	}


