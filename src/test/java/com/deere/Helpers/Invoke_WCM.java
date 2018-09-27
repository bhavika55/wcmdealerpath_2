package com.deere.Helpers;

import java.io.IOException;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Set;

import org.openqa.selenium.WebDriver;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import com.deere.PageFactory.Login_Page_POF;
import com.deers.alerts_WCM.Alert_WCM_POF;


public class Invoke_WCM extends BaseClass {
	
	
	WebDriver driver;
	/**
	 * This method is the first step of DealerPath suite which sets user
	 * credentials, initiate drivers and page elements
	 * 
	 * @author shrishail.baddi
	 * @createdAt 07-06-2018
	 * @throws IOException
	 * @throws Exception
	 * @modifyBy shrey.choudhary
	 * @modifyAt
	 */
	@BeforeClass
	public void systemConfigSetup() throws IOException, Exception {
		try {

			WCMInput.readWCMContentData();
			wcmInputdata=WCMInput.WCMInputValues;
			commonInputValues=WCMInput.InputValues;
			
			URL=commonInputValues.get("URL");
			strBrowserType=commonInputValues.get("Browser");
			strUserName=commonInputValues.get("Username");
			strPassword=commonInputValues.get("Password");
			
			
				BrowserFactory.initiateDriver();
				initPageElements();
				
			} catch (Exception e) {
				LogFactory.info(e.getMessage());
			} catch (Throwable e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * This method is use to invoke admin's login credentials then go to impersonate
	 * the dealer
	 * 
	 * @author shrey.choudhary
	 * @createdAt 07-06-2018
	 * @throws IOException
	 * @throws Exceptionss88593
	 * @modifyBy
	 * @modifyAt
	 * @throws Throwable
	 */
	@Test(priority=0)
	public static void invokeUserCredentials() throws Throwable {
		
		System.out.println("Verify Valid Login");
		
		WCMInput.readWCMContentData();
		wcmInputdata=WCMInput.WCMInputValues;
		commonInputValues=WCMInput.InputValues;
		
		URL=commonInputValues.get("URL");
		strBrowserType=commonInputValues.get("Browser");
		strUserName=commonInputValues.get("Username");
		strPassword=commonInputValues.get("Password");
		
		
		System.out.println("**Reading input from testdata excel**");
		Set<String> keys = wcmInputdata.keySet();
		 for(String key: keys)
		 	{
			HashMap<String, Object> InnerMap =	 wcmInputdata.get(key); 
			System.out.println("Library::"+InnerMap.get("Library")); 
			System.out.println("Departments::"+InnerMap.get("Departments")); 
			System.out.println("published date::"+InnerMap.get("Published Date"));
			System.out.println("number of contents to fetch::"+InnerMap.get("Number Of Rows to fetch"));
			
			}
		
		Login_Page_POF.setCredentials(strUserName, strPassword);

		if (Login_Page_POF.verifyUserLogin()) {
		
		/*	System.out.println("**Reading input from testdata excel**");
			Set<String> keys = wcmInputdata.keySet();
			 for(String key: keys) 
			 	{
				HashMap<String, Object> InnerMap =	 wcmInputdata.get(key); 
				System.out.println(InnerMap.get("Departments")); 
				}*/
			 
			
			System.out.println("Navigating to WCM page");
			Login_Page_POF.navigateToWCM();
			
		}
		
		else {
			
			System.out.println("Login for"+BaseClass.strUserName+"Failed");
		}
	}
	
	
	//read all WCM content for a region 	
	@Test(priority=1)
	public static void moveToAlert() throws Throwable{
		try {
		System.out.println("***Test case for WCM content verification***");
		
			Alert_WCM_POF.createWCMExcel();
		
			Alert_WCM_POF.navigateToRegion(alertRegion);
			
			//Alert_WCM_POF.readWCMAlertsAnnouncementsContent();
			
			Alert_WCM_POF.fetchDepartmentContents(department);
	         
		
		}
		catch(Exception e)
		{
			System.out.println("Error while navigating to alert section::"+ e.getMessage().toString());
			
		}
	} 
		
	
/*	@AfterClass
	public void closeDriver() {
		
		
		BaseClass.wbDriver.quit();
		System.out.println("all windows closed sucessfully");
	}*/
}