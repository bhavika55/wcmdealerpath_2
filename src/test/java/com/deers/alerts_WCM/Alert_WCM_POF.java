package com.deers.alerts_WCM;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import javax.security.auth.callback.ChoiceCallback;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import com.deere.Helpers.BaseClass;
import com.deere.Helpers.ValidationFactory;
import com.deere.Helpers.WaitFactory;
import com.steadystate.css.parser.selectors.SyntheticElementSelectorImpl;
import com.steadystate.css.util.ThrowCssExceptionErrorHandler;

public class Alert_WCM_POF extends BaseClass{

	static WebDriver alrtDriver;
	private static String filename = "";
	static XSSFWorkbook workbook = null;
	static XSSFSheet spreadsheet = null;
	static String alertName=null;
	
	private static XSSFWorkbook wcmbook;
	private static XSSFSheet wcmdataSheet;
	
	static int testcaseNumber=1;
	
	static String testCaseID="WCM_TC";
	public static List<Map<String,String>> finalResultforExcel = new ArrayList<>();
	
	
	
	public Alert_WCM_POF(WebDriver driver)
	{
		this.alrtDriver=driver;
		
	}
	
	
	
	@FindBy(how = How.XPATH, using = "//a[contains(@id,'pageSizeFiftywcmTable')]")
	public static WebElement bottomPagenumber;
	
	@FindBy(how = How.XPATH, using = "//table//tr[contains(@id,'wcmTable_')]")
	public static WebElement allAlerts;
	
	
	@FindBy(how = How.XPATH, using = "//td[contains(.,'Published') and @class='lotusMeta']")
	public static WebElement allPublishedAlerts;
	
	
	@FindBy(how = How.XPATH, using = "//h4[@role='presentation']//a[contains(.,'Content')]")
	public static WebElement contentSection;

	@FindBy(how = How.XPATH, using = "//li[@class='wcmBreadcrumbsElement']//a[contains(.,'Content')]")
	public static WebElement contentNavigtionSection;

	@FindBy(how = How.XPATH, using = "//h4[@role='presentation']//a[contains(.,'My DealerPath')]")
	public static WebElement myDealerPath;
	
	@FindBy(how = How.XPATH, using = "//h4[@role='presentation']//a[contains(.,'Announcements')]")
	public static WebElement announcementNavigation;
	
	@FindBy(how = How.XPATH, using = "//a[@id='close_controllable']")
	public static WebElement closeContent;
	
	
	@FindBy(how = How.XPATH, using = "//h4[@role='presentation']//a[contains(.,'Alerts')]")
	public static WebElement alertsSection;
	
	
	@FindBy(how = How.XPATH, using = "//table//tr[contains(@id,'wcmTable_')]")
	public static List<WebElement> totalAlerts;	
	
	
	@FindBy(how = How.XPATH, using = "//td[contains(.,'Published') and @class='lotusMeta']")
	public static List<WebElement> totalPublishedAlerts;
	
	@FindBy(how = How.XPATH, using = "//*[@id='content_template']")
	public static WebElement contentTypeOnPage;
	
	@FindBy(how = How.XPATH, using = "//*[@id='id_ctrl_titlecom.aptrix.pluto.content.Content']")
	public static WebElement titleOnPage;
	
	@FindBy(how = How.XPATH, using = "//*[@id='locationcom.aptrix.pluto.content.Content']")
	public static WebElement locationOnPage;
	
	@FindBy(how = How.XPATH, using = "//*[@id='breadcrumb_library']")
	public static WebElement libraryOnPage;	

	@FindBy(how = How.XPATH, using = "//label[.='China MRU-Country']/following::div[1]")
	public static WebElement mruChinaCountry;	

	@FindBy(how = How.XPATH, using = "//label[.='MRU-Country']/following::div[1]")
	public static WebElement mruCountry;	

	@FindBy(how = How.XPATH, using = "//label[.='Product Type']/following::div[1]")
	public static WebElement productTypeOnPage;	
	
	@FindBy(how = How.XPATH, using = "//label[.='China Product Type']/following::div[1]")
	public static WebElement chinaProductTypeOnPage;	
	
	@FindBy(how = How.XPATH, using = "//label[.='Department']/following::div[1]")
	public static WebElement departmentOnPage;

	@FindBy(how = How.XPATH, using = "//label[.='Copy Department']/following::div[1]")
	public static WebElement copyDepartmentOnPage;
	
	@FindBy(how = How.XPATH, using = "//a[.='Edit']")
	public static WebElement editContent;
	
	@FindBy(how = How.XPATH, using = "//a[.='Read']")
	public static WebElement readContent;
	
	@FindBy(how = How.XPATH, using = "//label[.='Site Area Template:']/following::div[1]/a")
	public static WebElement siteArea;

	@FindBy(how = How.XPATH, using = "//a[contains(.,'Filter')]")
	public static WebElement filter;
	
	@FindBy(how = How.XPATH, using = "//a[contains(@id,'OkBtn')]")
	public static WebElement filterOk;
	
	@FindBy(how = How.XPATH, using = "//*[@id='breadcrumb_library']")
	public static WebElement checkForGlobalContent;
	
	@FindBy(how = How.XPATH, using = "//label[contains(.,'MRU-Country')]/following::div[1]")
	public static List<WebElement> globalCountries;
	
	@FindBy(how = How.XPATH, using = "//label[contains(.,'Product Type')]/following::div[1]")
	public static List<WebElement> globalProductType;
	
	@FindBy(how = How.XPATH, using = "//label[contains(.,'Type a URL:')]/following::div[1]/span")
	public static WebElement linkText;
	
	@FindBy(how = How.XPATH, using = "//label[contains(.,'File:')]/following::div[1]//span")
	public static WebElement documentText;
	
	@FindBy(how = How.XPATH, using = "//tr[contains(@id,'wcmTable_')]//td[2]//img[2]")
	public static List<WebElement> allChildren;
	
	@FindBy(how = How.XPATH, using = "label[contains(.,'Rich Text')]")
	public static WebElement richText;
	
	
	@FindBy(how = How.XPATH, using = "//*[@id='Link_link']")
	public static WebElement webContentLink;
	
	
	public static void clickOnBottomNumber() throws Throwable
	{
		
		try
		{
			Thread.sleep(3000);
			
				bottomPagenumber.click();
				System.out.println("Page Number link clicked sucessfully");
				
			
			
		}
		catch(Exception e)
		{
			e.printStackTrace();
			System.out.println("Error while clicking the expand number link:"+e.getMessage().toString());
		}
		
	}
	
		
	public static void readWCMAlertsAnnouncementsContent() throws Throwable
	{
				
		System.out.println("**Fetching individual alerts content**");		
		try
		{
		
			System.out.println("***navigating to alerts section***");
			
			System.out.println("***clciking content section");
			
			contentSection.click();
			
			System.out.println("***clciking Alerts section");
			
			alertsSection.click();
		
			applyFilterForStatus();
			
			moveInsideWCMContents("Alerts");
			
			System.out.println("***All alerts data fetched***");
			System.out.println("*Now navigating to Announcement section*");
			
			contentNavigtionSection.click();
			
			myDealerPath.click();
			
			announcementNavigation.click();
			
			Thread.sleep(3000);
		
			moveInsideWCMContents("Announcement");
			
		}
		
	

	catch(Exception e)
	{
		System.out.println("Link not clicked "+e.getMessage().toString());
	}
		
	
	}
	
	
	
	
	public static void moveInsideWCMContents(String wcmsection) throws Throwable
	{
		Map<String, String> wcmKeyvalue=new HashMap<String, String>();
		List<String> alertsLList;
		
	 try {
		 
		 System.out.println("***fetching WCM contents for "+wcmsection+" ***");
		 List<WebElement> publishedAlerts=ValidationFactory.getElementsIfPresent(By.xpath("//td[contains(.,'Published') and @class='lotusMeta']/preceding-sibling::td[1]//a[not(contains(@title,'View children'))]"));
			
			Iterator<WebElement> iter = publishedAlerts.iterator();
			
			alertsLList = new ArrayList<String>();
			
			while(iter.hasNext()) 
			{
				WebElement w = iter.next();
				alertsLList.add(w.getText());
		 	}
		
			System.out.println("Total Published "+wcmsection+" in list are:"+alertsLList.size());
			
			 for(int i=0;i<alertsLList.size();i++) 
			 	{
				 System.out.println("Fetching content for "+wcmsection+" :"+alertsLList.get(i));
				 WebElement alert1=alrtDriver.findElement(By.xpath("//a[contains(.,'"+alertsLList.get(i)+"')]"));
				
			 	alert1.click();
				 	
					String wcmTCID=testCaseID+testcaseNumber;
					wcmKeyvalue.put("WCMSection", wcmsection);
			  		wcmKeyvalue.put("Test Case ID", wcmTCID);
			    	
			  		writeWCMToExcel(wcmKeyvalue);
			  		writeWCMHeaderContentFinalToExcel();
			  		testcaseNumber++;
			    	closeContent.click();
	 										
			 	}
	 	}
			 catch(Exception e)
			 {
				 
				 System.out.println("Error while writing contents for "+ wcmsection+" " +e.getMessage().toString());
			 }
	 }
		
		
	
	
	
	public static void navigateToRegion(String alertRegionLanguage) throws Throwable
	{
					
	try
	{
		System.out.println("**Inside Alert navigation method**");
		
		if (alertRegionLanguage != null) {

			bottomPagenumber.click();
			alrtDriver.findElement(By.xpath("//a[contains(.,'"+alertRegionLanguage+"')]")).click();
			
		}
		
	}
		catch(Exception e)
		{
		e.printStackTrace();
		System.out.println("Couldn't navigate to alert section "+e.getMessage().toString());
		}
		
	
	}
	
	
	
	public static void applyFilterForStatus() throws Throwable{
		
		try {
			
			filter.click();
					
			String clickFilter="//*[@id='ibm_wcm_widget_filter_FilterField_0_menuLink']";
			WebElement filterclicked= WaitFactory.explicitWaitByXpath(clickFilter);
		
			filterclicked.click();
			Thread.sleep(2000);
		
			String selectingStatus="//td[contains(@id,'_STATUS_text')]";
			
			WebElement statusSelect= WaitFactory.explicitWaitByXpath(selectingStatus);
		
			statusSelect.click();
			
			Thread.sleep(1000);
			filterOk.click();
		
			Thread.sleep(2000);
			
		}
		catch(Exception e)
		{
				System.out.println("Unable to apply filter "+e.getMessage().toString());
		}
		
	}
	


public static void applyFilterForDate() throws Throwable{
		
		try {
			System.out.println("***Applying filter for Date***");
			
			filter.click();
					
			alrtDriver.findElement(By.xpath("//a[.='Add a filter']")).click();
			
			
			String clickFilter="//*[@id='ibm_wcm_widget_filter_FilterField_1_menuLink']";
			WebElement filterclicked= WaitFactory.explicitWaitByXpath(clickFilter);
		
			filterclicked.click();
			Thread.sleep(2000);
		
			String selectingLastSaveddate="//td[contains(@id,'_MODIFIED_DATE_text')]";
			
			WebElement lastSavedDateSelect= WaitFactory.explicitWaitByXpath(selectingLastSaveddate);
		
			lastSavedDateSelect.click();
			
			
			String selectingAfterDropDown="//a[@id='ibm_wcm_widget_filter_DateFilter_0_condition']";
			
			WebElement afterDropDownSelect= WaitFactory.explicitWaitByXpath(selectingAfterDropDown);
		
			afterDropDownSelect.click();
			
			

			String selectingAfter="//td[contains(@id,'_conditionAFTER_text')]";
			
			WebElement afterSelect= WaitFactory.explicitWaitByXpath(selectingAfter);
		
			afterSelect.click();
			
			
			
			String dateDropDown="//div[@id='widget_ibm_wcm_widget_filter_DateFilter_1_date1']//input[contains(@class,'dijitArrowButtonInner')]";
			WebElement dateDropDownSelect= WaitFactory.explicitWaitByXpath(dateDropDown);
			dateDropDownSelect.click();
			
			
			//selecting the date code goes here
			String publishedDate="20-September-2018";
			String []dateArray=publishedDate.split("-");
			String date=dateArray[0];
			String month=dateArray[1];
			String year=dateArray[2];
			
			DateTimeFormatter dateFormat = DateTimeFormatter.ofPattern("yyyy/MM/dd");
			LocalDate localDate = LocalDate.now();
			System.out.println(dateFormat.format(localDate));
			
			
	        String strDate = dateFormat.format(localDate);

	        String currentdate[]=strDate.split("/");
			String currentYear=currentdate[0]; 
	        
			//selecting year
			
			
			if(alrtDriver.findElement(By.xpath("//span[.='"+year+"']")).isDisplayed())
			{
				String selectYear="//span[.='"+year+"']";
				WebElement YearSelect= WaitFactory.explicitWaitByXpath(selectYear);
				YearSelect.click();
				
			}
			
			else if(Integer.parseInt(year)<Integer.parseInt(currentYear) && alrtDriver.findElement(By.xpath("//span[.='"+year+"']")).isDisplayed())
	        {
	        	alrtDriver.findElement(By.xpath("//span[.='"+year+"']")).click();
	        }
	        
			else
			{
				
				
			}
			
			String selectYear="//span[.='"+year+"']";
			WebElement YearSelect= WaitFactory.explicitWaitByXpath(selectYear);
			YearSelect.click();
			
			
			//selecting Month
			
			
			//selecting day
			
			alrtDriver.findElement(By.xpath("//span[.='"+date+"']")).click();
			
			
			Thread.sleep(1000);
			filterOk.click();
		
			Thread.sleep(2000);
			
		}
		catch(Exception e)
		{
				System.out.println("Unable to apply filter "+e.getMessage().toString());
		}
		
	}
	
	
	public static List<WebElement> identifyAlllanguages(String alertRegion)throws Throwable {

		
		try {
			
			
			bottomPagenumber.click();
			
			Thread.sleep(5000);
			List<WebElement> allLanguages= ValidationFactory.getElementsIfPresent(By.xpath("//a[contains(.,'"+alertRegion+"_CONTENT_')]"));
			
			
			System.out.println("***Total Languages available for the Region "+alertRegion+" are "+allLanguages.size());
			
			Iterator<WebElement> iter = allLanguages.iterator();

					
			while(iter.hasNext()) {
			    
					WebElement we = iter.next();

			    	System.out.println("**Region available is: "+we.getText());
			         
						}
			return allLanguages;
			
			}
		
		
		catch(Exception e)
		{
			
			System.out.println("elements not found"+e.getMessage().toString());
			
		}
		return null;
	}
	
	
	public static void writeWCMHeaderContentFinalToExcel() throws Throwable
	{
		
		try
		{
			System.out.println("***Writing final content into WCM Excel***");
			writeWCMHeader(filename, BaseClass.headerList);

			writeWCMRow(filename, BaseClass.finalResultforExcel);

		}
		catch(Exception e)
		{
			
			System.out.println("error while writing WCM content excel"+e.getMessage().toString());
		}
	}

	
	public static void createWCMExcel() throws Throwable{
		
				try {

			DateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy_HH-mm-ss");
			Date date = new Date();
			filename = wcmDataOutputPath + dateFormat.format(date) + ".xlsx";
				
			System.out.println("Excel File to be created is :: "+filename);
			
			System.out.println("**WCM Excel created successsfully**");
			
			System.out.println("**Now adding WCM content Headers into List**");
			 BaseClass.headerList= new  ArrayList<String>();
			   
			  BaseClass.headerList.add("Test Case ID"); 
			  BaseClass.headerList.add("EXECUTE");
			  BaseClass.headerList.add("URL");
			  BaseClass.headerList.add("Library");
			  BaseClass.headerList.add("Multilingual");
			  BaseClass.headerList.add("DepartmentName");
			  BaseClass.headerList.add("2ndLevel");
			  BaseClass.headerList.add("3rdLevelIndexPage");
			  BaseClass.headerList.add("3rdLevelIndexPageCategories");
			  BaseClass.headerList.add("3rdLevelIndexPageNestedCategories");
			  BaseClass.headerList.add("3rdLevelLandingPage");
			  BaseClass.headerList.add("3rdLevelChildIndexPage");
			  BaseClass.headerList.add("3rdLevelChildIndexPageCategories");
			  BaseClass.headerList.add("3rdLevelChildIndexPageNestedCategories");
			  BaseClass.headerList.add("3rdLevelGrandChildIndexPage");
			  BaseClass.headerList.add("3rdLevelGrandChildIndexPageCategories");
			  BaseClass.headerList.add("3rdLevelGrandChildIndexPageNestedCategories");
			  BaseClass.headerList.add("3rdLevelFolder");
			  BaseClass.headerList.add("4thLevelIndexPage");
			  BaseClass.headerList.add("4thLevelIndexPageCategories");
			  BaseClass.headerList.add("4thLevelIndexPageNestedCategories");
			  BaseClass.headerList.add("4thLevelLandingPage");
			  BaseClass.headerList.add("4thLevelChildIndexPage");
			  BaseClass.headerList.add("4thLevelChildIndexPageCategories");
			  BaseClass.headerList.add("4thLevelChildIndexPageNestedCategories");
			  BaseClass.headerList.add("4thLevelGrandChildIndexPage");
			  BaseClass.headerList.add("4thLevelGrandChildIndexPageCategories");
			  BaseClass.headerList.add("4thLevelGrandChildIndexPageNestedCategories");
			  BaseClass.headerList.add("ContentType");
			  BaseClass.headerList.add("IndexPageContentType");
			  BaseClass.headerList.add("Title");
			  BaseClass.headerList.add("Keywords");
			  
			  BaseClass.headerList.add("DocPath");
			  BaseClass.headerList.add("Link");
			  BaseClass.headerList.add("Description");
			  BaseClass.headerList.add("ReleaseDate");
			  BaseClass.headerList.add("Column4");   
			  BaseClass.headerList.add("Column5");
			  BaseClass.headerList.add("MRU-Country");
			  BaseClass.headerList.add("ProductType");
			  BaseClass.headerList.add("DealerType (Main/Sub)");
			  BaseClass.headerList.add("Index_Page_Template");
			  BaseClass.headerList.add("Index_Page_Template_Label");
			  BaseClass.headerList.add("RACFGroups");
			  BaseClass.headerList.add("CopyToDepartment");
			  BaseClass.headerList.add("Comments");
			 
					      
		}
		catch(Exception e)
		{
			
			System.out.println("error while creating WCM hedaer content list"+e.getMessage().toString());
		}
		
	}
	
	
	
	public static void writeWCMToExcel(Map<String , String> valuesToWrite ) throws Throwable{
			
		
				String contentType=null;
				String title=null;
				String countryTitle =null;
				String location=null;
				String library=null;
				String department=null;
				String copyToDepartment=null;
				String indexPageContentType=null;
				String link=null;
				String document=null;
				String prodTypeFinal=null;
		try {


			contentType=contentTypeOnPage.getText();	
			String[] cType=contentType.split("/");
			String conType =  cType[cType.length-1].trim();
	
			title=titleOnPage.getText();	
			location=locationOnPage.getText();
			library=libraryOnPage.getText();
			
			if(!conType.equals("AT-Default"))
			{
				if(checkForGlobalContent.getText().contains("GLOBAL_CONTENT"))
				{
					System.out.println("**This is a global content**");
					String libraryRegions[]=alertRegion.split("_");
					String library1=libraryRegions[0];
					
					String countries=alrtDriver.findElement(By.xpath("//label[.='"+library1+"-MRU-Country']/following::div[1]/span")).getText();
					countryTitle=fetchCountriesList(countries);
						
					String products=alrtDriver.findElement(By.xpath("//label[.='"+library1+"-Product Type']/following::div[1]/span")).getText();;
					prodTypeFinal=fetchProductsList(products);
				}
				
				else {
					
				Actions actions = new Actions(alrtDriver);
				actions.moveToElement(mruCountry);
				actions.build().perform();
				
				String countries1=mruCountry.getText();
				countryTitle=fetchCountriesList(countries1);
				
				prodTypeFinal=fetchProductsList(productTypeOnPage.getText());
		
				}
			}
		   for(Map.Entry<String, String> entry : valuesToWrite.entrySet())
		   {
			   System.out.println("wcmSection is:"+entry.getValue());
			   System.out.println(entry);
			  
		    	if(entry.getKey().contains("WCMSection"))
		    	{
		    		System.out.println("Inside wcm");
		    		String wcmsection=valuesToWrite.get("WCMSection");
		    	
		    		
		    		if(wcmsection.equals("Alerts"))
		    		{
		    			System.out.println("Content type is alerts");
		    			break;
		    		}
		    		else if(wcmsection.equals("Announcement"))
		    		 {	
		    			 if(conType.equals("AT-Announcement"))
		    			 	{
		    				 String departmentType=departmentOnPage.getText();
		    				 String[] deptType=departmentType.split("/");
		    				department =  deptType[deptType.length-1].trim();
		    				copyToDepartment=fetchProductsList(copyDepartmentOnPage.getText());
		    				break;
		    			 	}
		    		 }
		        
		    		else if(wcmsection.equals("Announcement for Departement"))
		        		{
		        		System.out.println("Inside announcement for department");
		    			/*String checkForDuplicateAnnouncement=alrtDriver.findElement(By.xpath("//li/a[contains(.,'Announcements')]/preceding::li[2]//a")).getText();
		        			if(checkForDuplicateAnnouncement.equals("My DealerPath"))
		        			{
								System.out.println("Announcement: "+title+" already exists");
		        				break;
		        			}*/
		        	
		        			if(conType.equals("AT-Announcement"))
		        			{ 
		        				System.out.println("Inside AT-Announcement");
		        				String departmentType=departmentOnPage.getText();
		        				String[] deptType=departmentType.split("/");
		        				department =  deptType[deptType.length-1].trim();
		        				System.out.println("Department copied");
		        				copyToDepartment=fetchProductsList(copyDepartmentOnPage.getText());
		        				System.out.println("copy deprtment copied");
		        				break;
		    				}
		        			else 
		        			{
		        			System.out.println("inside other content");
		        			department=alrtDriver.findElement(By.id("breadcrumb_item_1")).getText();
		        				break;
		        			}
		        
		        		}
		    	
		    } ///this is for alerts and announcements
		    	
		    	else
		    	{
		    		//System.out.println("inside not wcm section");
		    		
		    		if(conType.contains("Default"))
		    		{
		    			department=alrtDriver.findElement(By.id("breadcrumb_item_1")).getText();
        				break;
		    		}
		    		
		    		else if(conType.equals("AT-Link") || conType.contains("AT-Link_Devl"))	
		    		{
		    		if(ValidationFactory.isElementPresent(linkText))
		    			{
		    				link=linkText.getText();
		    				indexPageContentType="Link";
		    				break;
		    				
		    			}
		    			
		    			else if(ValidationFactory.isElementPresent(alrtDriver.findElement(By.xpath("//*[@id='Link_link']"))) && !((alrtDriver.findElement(By.xpath("//*[@id='Link_link']"))).getText().contains("None")))
		    			{
		    				link=alrtDriver.findElement(By.xpath("//*[@id='Link_link']")).getText();
		    				indexPageContentType="Link";
		    				break;
		    				
		    			}	
			    		
		    		}
		    		
		    		
		    	
		    		else if(conType.equals("AT-Document") || conType.contains("AT-Embedded Document"))
		    		{
		    		document=documentText.getText();
		    		indexPageContentType="Document";
		    		break;
		    		}
		    		
		    		else if(ValidationFactory.isElementPresent(richText) && ValidationFactory.isElementPresent(By.xpath("body[@spellcheck='true']//*")))
		    		{
		    			indexPageContentType="Rich-Text";
		    			break;
		    		}
		    				
		    		else 
		    		{		    			
		    		System.out.println("inside content not verififed");
		    		String contentIs=checkForLinkRichTextDocument();
		    		System.out.println("Content is "+contentIs);
		    		String fetchFirstLast[]=contentIs.split("/");
		    		String contentIsFinal=fetchFirstLast[fetchFirstLast.length-1].trim();
		    			
		    			if(contentIs.contains("Link")) 
		    			{
		    				link=contentIsFinal;
		    				indexPageContentType="Link";
		    				break;
		    				
		    			}
		    			else if(contentIs.contains("Document"))
		    			{
		    				document=contentIsFinal;
		    				indexPageContentType="Document";
		    				break;
		    			}
		    			
		    		}
		    	
		    	}
		    		
		    	}
		    	
		   
		    System.out.println("Now mapping the fetched WCM content into excel");
		    
		   Map<String, String> finalkeyValueWCM; 
		    finalkeyValueWCM = new HashMap<String, String>();
		    
		     // Adding elements to the recently created HashMap
		    finalkeyValueWCM.put("ContentType", conType);
		    finalkeyValueWCM.put("Title", title);
		    finalkeyValueWCM.put("Location", location);
		    finalkeyValueWCM.put("Library", library);
		    finalkeyValueWCM.put("DepartmentName", department);
		    finalkeyValueWCM.put("CopyToDepartment", copyToDepartment);
		    finalkeyValueWCM.put("MRU-Country", countryTitle);
		    finalkeyValueWCM.put("ProductType", prodTypeFinal);
		    finalkeyValueWCM.put("IndexPageContentType", indexPageContentType);
		    finalkeyValueWCM.put("Link", link);
		    finalkeyValueWCM.put("DocPath", document);		    
		    
		    
		    // Copying one HashMap "hmap" to another HashMap "hmap2"
		    finalkeyValueWCM.putAll(valuesToWrite);
		     
		    excelOutput(finalkeyValueWCM);
		     
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			System.out.println("Error while fetching content from "+e.getMessage().toString());


		}
	}
	
		
	
	
	
	private static String fetchProductsList(String products) throws Throwable{
		// TODO Auto-generated method stub
		String productTypeTitle=null;
		
		try
		{
			
				
				String[] prodTypeList=products.split(",");
				
				if(prodTypeList.length>=2)
				{
				for(int n=0;n<=prodTypeList.length-1;n++)
				{
					
					String []diffrentProductType=prodTypeList[n].split("/");
					String listOfProdType=diffrentProductType[diffrentProductType.length-1].trim();
					
					productTypeTitle=productTypeTitle+","+listOfProdType.trim();
				}
				
				if(productTypeTitle.startsWith("null"))
					productTypeTitle=productTypeTitle.substring(5);
				
				}
				return productTypeTitle;
			
			
		}
		catch(Exception e)
		{
			System.out.println("Error while fetchin product types from ceontent"+e.getMessage().toString());
			
		}
		return null;
	}


	private static String checkForLinkRichTextDocument() throws Throwable{
		
		try
		{ 
			
			System.out.println("verfying the content for Link, Rich-Text or Document");
					
			if(ValidationFactory.isElementPresent(linkText) && !(linkText.getText().contains("None")))
			{
				//if((linkText.getText().contains("None"))) 
				return "Link/"+linkText.getText();
				
			}
			
			else if(ValidationFactory.isElementPresent(documentText) && !(documentText.getText().contains("None")))
			{
				
				return "Document/"+documentText.getText();
			}
			
						
			else if(ValidationFactory.isElementPresent(richText) && ValidationFactory.isElementPresent(alrtDriver.findElement(By.xpath("body[@spellcheck='true']//*"))))
			{return "Rich-Text";}
			else if(ValidationFactory.isElementPresent(webContentLink) && !(webContentLink.getText().contains("None")))
			{
				
				return "Link/"+webContentLink.getText();
				
			}
			
			
			
		}
		
		catch(Exception e)
		{
			
			System.out.println("Error while determining the type of content "+e.getMessage().toString());
		}
		
		
		return null;
	}


	public static void excelOutput(Map<String, String> wcmContentToWrite) throws Throwable {
		
		
		System.out.println("***mapping alerts contents into List***");
		try {

		 //System.out.println("Final WCM content as Key Value pair" +wcmContentToWrite);
		  BaseClass.excelList = new LinkedHashMap<String,String>();
		  
		  for(Map.Entry<String, String> entry : wcmContentToWrite.entrySet())
		  {
			 BaseClass.excelList.put(entry.getKey(), entry.getValue());
		  }
				  
		  System.out.println("Key Value hashmap: "+BaseClass.excelList);
		     
		  BaseClass.finalResultforExcel.add(BaseClass.excelList);
			}
		
		catch(Exception e) 
		{
		System.out.println("Error while mapping content to list "+e.getMessage().toString());
		}

	}
	
	
	
	
	
	public static String writeWCMHeader(String fileName, List<String> headerList) throws IOException {
		try
		{
		FileOutputStream fos = new FileOutputStream(new File(fileName));
		XSSFWorkbook book = new XSSFWorkbook();
		
		XSSFSheet sheet;
		
		sheet = book.createSheet("WCM Content");
				
		Row row = sheet.createRow(0);

		int cellNumber = 0;
		Font font = book.createFont();
		//font.setBold(true);
		font.setFontHeightInPoints((short) 9);
		font.setColor(IndexedColors.DARK_YELLOW.getIndex());
		font.setBold(true);
		
		CellStyle cellStyle1 = book.createCellStyle();

		for (String header : headerList) {
			cellStyle1.setFillForegroundColor(IndexedColors.GREEN.getIndex());
			cellStyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			cellStyle1.setFont(font);
			
			Cell cell = row.createCell(cellNumber++);
			
	
			cell.setCellValue(header);
			cell.setCellStyle(cellStyle1);
			
			sheet.autoSizeColumn(cellNumber);

		}
		book.write(fos);
		book.close();
		fos.close();
		
		return fileName;
		
		
		}
		catch(Exception e)
		{
			e.printStackTrace();
			System.out.println("Error while creating excel with Hedaer data"+e.getMessage().toString());
			
			
		}
		return fileName;
	}
	
	
	public static void writeWCMRow(String fileName, List<Map<String, String>> rowList)
			throws IOException, InvalidFormatException, Throwable {
		
		try {
			int rowNum;
			
		File oFile = new File(fileName);
		FileInputStream input = new FileInputStream(oFile);
		wcmbook = new XSSFWorkbook(input);

		wcmdataSheet = wcmbook.getSheet("WCM Content");	
		
		rowNum = wcmdataSheet.getLastRowNum();
		CellStyle style = wcmbook.createCellStyle();// *
		//Font font = wcmbook.createFont();// *

		for (int i = 0; i < rowList.size(); i++) {
			Map<String, String> map = new HashMap<String, String>();
			map = rowList.get(i);
			XSSFRow row = wcmdataSheet.createRow(++rowNum);
			
			
			XSSFRow headerRow=wcmdataSheet.getRow(0);
			
			short minColIx = headerRow.getFirstCellNum(); //get the first column index for a row
			short maxColIx = headerRow.getLastCellNum(); //get the last column index for a row

			for(short colIx=minColIx; colIx<maxColIx; colIx++) { //loop from first to last column index
				   
				XSSFCell cell = headerRow.getCell(colIx); //get the cell
				 
				   //add the cell contents (name of column) and cell index to the map   
				 String headerName= cell.getStringCellValue();
				 
				 int headerIndex=cell.getColumnIndex();
				 
				 XSSFCell cellToWrite = row.createCell(headerIndex);
				
			
			for (Map.Entry<String, String> entry : map.entrySet()) {
				if(map.containsKey(headerName))
				{
					String cellVal1= map.get(headerName);
					
					if (cellVal1 instanceof String) {
						cellToWrite.setCellValue((String) cellVal1);
							}
					
					
					else {
						cellToWrite.setCellValue("NA");
						}
					
				}
				
				else
				{
					
					cellToWrite.setCellValue("NA");
				}
			
				wcmdataSheet.autoSizeColumn(headerIndex);
				
				style.setWrapText(true);
				cell.setCellStyle(style);
			}// end of for loop for writing content into correct cell
			
			
			}//End of For loop for matching Excel header with map's key
			
			
		}//end of outer for loop writing all rows to excel
		
		
		input.close();
		FileOutputStream fos = new FileOutputStream(oFile);
		wcmbook.write(fos);
		wcmbook.close();
		fos.close();
		
	}
		catch(Exception e)
		{
			System.out.println("Error While wrting row content"+e.getMessage().toString());
			
		}
	}


	public static void fetchDepartmentContents(String departmentName) throws Throwable{
		
		
		HashMap<String, String> wcmKeyValue1= new HashMap<String, String>();
		
		System.out.println("**Now fetching Department wise data**");
		
		try
		{
			
			contentSection.click();
			
			applyFilterForStatus();
			
			myDealerPath.click();
						
			alrtDriver.findElement(By.xpath("//a[contains(.,'"+departmentName+"')]")).click();
			
			System.out.println("fetching Announcements for department: "+departmentName);
			
			alrtDriver.findElement(By.xpath("//a[contains(.,'Announcements')]")).click();
			
			Thread.sleep(1000);
			
			moveInsideWCMContents("Announcement for Departement");
			
			System.out.println("Announcements for Department "+departmentName+" has been fetched proerly");
			
			alrtDriver.findElement(By.xpath("//a[contains(.,'"+departmentName+"')]")).click();
			
		//////////////Department//////////////////////////////////	
			
			// fetching all subChildrens 
			List<WebElement> allImagesUnderDepartment= allChildren;
			List<String> subDepartments = new ArrayList<String>();//contains all subDepartments under Business Admin & HR
			
			for(int i=3;i<=allImagesUnderDepartment.size();i++)
			{	
				String FecthDeptsimageTitle=alrtDriver.findElement(By.xpath("//tr["+i+"]//td[2]//img[2]")).getAttribute("title");
				if(FecthDeptsimageTitle.contains("View children"))
				{
					String DeptChildrenName=alrtDriver.findElement(By.xpath("//tr["+i+"]//td[2]//img[2]/following::td[1]//a")).getText();
					subDepartments.add(DeptChildrenName);
				}
			}

			
//Now writing departments(Business Admin & HR) with children into excel e.g (subDeptsUnderDeptName)Optimization,Business Management etc. and its children(Links, documents and Index pages)
	
	//DepthasChildren contains all SAT-Index pages, Landing pages, SAT-Folders folders etc
			
//////////////Sub Departments////////////////			
			System.out.println("TOTAL SubDepartments under department: "+departmentName+" are::"+subDepartments.size());
			for(int sd=1;sd<=subDepartments.size();sd++)
			{
				
				List<String> SAT_Index_pages=new ArrayList<String>();;
				List<String> SAT_Folders=new ArrayList<String>();;
				List<String> SAT_LandingPages= new ArrayList<String>();
				List<String> SAT_Table_Index_pages=new ArrayList<String>();
				
				
				 String subDeptsUnderDeptName=subDepartments.get(sd); //Optimization (Sub Department)
		    	
				 System.out.println("Fetching content for SubDepartment: "+subDeptsUnderDeptName);//Optimization
				
				 WebElement subDepartment=alrtDriver.findElement(By.xpath("//a[.='"+subDeptsUnderDeptName+"' and starts-with(@title,'View children')]"));
			
				 subDepartment.click(); //Optimization or Business Management Clicked
				 
			//////////Inside Sub Department "Optimization"////////////////	 
				 Thread.sleep(1000);
			    	
//now fetching all contents under sub department e.g: Optimization's children(index pages and rest of the contents)
					
				 List<WebElement> allSubDeptChildrenImages= allChildren;					
					
			//adding contents under Sub Departments(Optimization) into different lists		
					
					List<String> SubDeptHasChildren = new ArrayList<String>();
					List<String> SubDeptLinkPortlets = new ArrayList<String>();
					
					
					for(int sdc=1;sdc<=allSubDeptChildrenImages.size();sdc++)
					{	
						String subDeptImageTitle=alrtDriver.findElement(By.xpath("//tr["+sdc+"]//td[2]//img[2]")).getAttribute("title");
						if(subDeptImageTitle.contains("View children"))
						{
							String SubDeptChildName=alrtDriver.findElement(By.xpath("//tr["+sdc+"]//td[2]//img[2]/following::td[1]//a")).getText();
							SubDeptHasChildren.add(SubDeptChildName);
						}
						else 
						{
							String SubDeptNoChildName=alrtDriver.findElement(By.xpath("//tr["+sdc+"]//td[2]//img[2]/following::td[1]//a")).getText();
							SubDeptLinkPortlets.add(SubDeptNoChildName);
						}
					}
					
					System.out.println("Total index like Childrens of SubDepartment:"+subDeptsUnderDeptName+" are "+SubDeptHasChildren.size());
					
			
				//Fetching Link Portlets for Sub Department Optimization in excel(Link,documents,Rich text)
					
						for(int k=0;k<numberOfContentsToFetch(SubDeptLinkPortlets);k++)
						{
						 String SubDeptLinkPortlet=SubDeptLinkPortlets.get(k);
					     System.out.println("Fetching content for "+SubDeptLinkPortlet);
						 WebElement subDeptlinkPortletElement=alrtDriver.findElement(By.xpath("//a[.='"+SubDeptLinkPortlet+"' and not(contains(@title, 'View children'))]"));
						 
						 subDeptlinkPortletElement.click();
						 String wcmTCID=testCaseID+testcaseNumber;
					    	wcmKeyValue1.put("Test Case ID",wcmTCID);
					    	wcmKeyValue1.put("DepartmentName",departmentName);
					    	wcmKeyValue1.put("2ndLevel",subDeptsUnderDeptName);
					    	
					    	   writeWCMToExcel(wcmKeyValue1);
					    	   writeWCMHeaderContentFinalToExcel();
					    	   testcaseNumber++;
					    	   closeContent.click();
						} 			//TILL THIS POINT CODE WORKING FINE//
					
					
//creating List of content type for Sub department e.g:Index pages, Tables,Folders, Landing Pages
					
					
					for(int m=0;m<SubDeptHasChildren.size();m++)
					{
							String indexchildName=SubDeptHasChildren.get(m);  //Business Continuation 
							String contentType= checkContentType(indexchildName);
						
						if(contentType.contains("SAT-Folder site Area"))
						{
							SAT_LandingPages.add(indexchildName);
						}
						else if(contentType.contains("Folder"))
						{
							SAT_Folders.add(indexchildName);
						}
						else if(contentType.contains("SAT-Index Page"))
						{	
							SAT_Index_pages.add(indexchildName);
						}
						else if(contentType.contains("SAT-Table Index Page"))
						{
							SAT_Table_Index_pages.add(indexchildName);
						}
						
					} //end of loop for fetching total tree icon contents type for Sub Department Optimization under Business Admoin & HR
						
					
					System.out.println("Total SAT_Index pages under:"+subDeptsUnderDeptName+" are "+SAT_Index_pages.size());
					System.out.println("Total SAT_Landing pages under: "+subDeptsUnderDeptName+" are "+SAT_LandingPages.size());
					System.out.println("Total SAT_Folders under: "+subDeptsUnderDeptName+" are "+SAT_Folders.size());
					System.out.println("Total SAT_Tables under: "+subDeptsUnderDeptName+" are "+SAT_Table_Index_pages.size());
				
					
				///LOGIC FOR FETCHING CONTENT FOR ALL TABLES	
					Map<String,String> deprtmentTables=new HashMap<String,String>();
					
					//for(int rt=0;rt<SAT_Table_Index_pages.size();rt++)
					for(int rt=0;rt<numberOfContentsToFetch(SAT_Table_Index_pages);rt++)
					{
						System.out.println("Table :"+SAT_Table_Index_pages.get(rt));
						deprtmentTables=fetchTablesContent(SAT_Table_Index_pages.get(rt));// can contain Link Portlets and Child index pages
						//departmentName,subDeptsUnderDeptName
					
						String wcmTCID=testCaseID+testcaseNumber;
						
						wcmKeyValue1.put("Test Case ID",wcmTCID);
				    	wcmKeyValue1.put("DepartmentName",departmentName);
				    	wcmKeyValue1.put("2ndLevel",subDeptsUnderDeptName);
				    	wcmKeyValue1.put("3rdLevelIndexPage",deprtmentTables.get("title"));
						wcmKeyValue1.putAll(deprtmentTables);
						
						excelOutput(wcmKeyValue1);
	      		        writeWCMHeaderContentFinalToExcel();
				        testcaseNumber++;
					
					}				

				//CALLING FUNCTION FOR FETCHING CONTENT FOR ALL INDEX PAGES 
					fetchContentTillGrandChild(SAT_Index_pages,departmentName,subDeptsUnderDeptName);

					alrtDriver.findElement(By.xpath("//a[.='"+departmentName+"']")).click();
					
					//fetchContentsForlandingPages(SAT_LandingPages,departmentName,subDeptsUnderDeptName);// can contain Tables, Link Portlets and Index pages
					//fetchContentsForFolders(SAT_Folders,departmentName,subDeptsUnderDeptName);// can contain Tables, Link Portlets and Index pages, landing pages
					
					}
				
			
		}
		catch(Exception e)
		{
			
			System.out.println("Error while fetching Department wise wcm Content for department: "+departmentName+" "+e.getMessage().toString());
		}
		
	} /// END OF fetchDepartmentContents method
	
	
	public static Map<String,String> fetchTablesContent(String Tablename)
	{
		String tableName=null;
		String totalCoulmns=null;
		
		Map<String,String> dataToFetch=new HashMap<String,String>();
		String fetchAttribute=alrtDriver.findElement(By.xpath("//a[.='"+Tablename+"' and contains(@title,'View children')]")).getAttribute("id");
		
		String[] checkboxNumber=fetchAttribute.split("_");
		String checkboxToClick =  checkboxNumber[checkboxNumber.length-1].trim();
		
		alrtDriver.findElement(By.xpath("//input[@value='"+checkboxToClick+"']")).click();
		
		editContent.click();
		//readContent.click();
		
		tableName=alrtDriver.findElement(By.xpath("//input[@id='id_ctrl_name']")).getAttribute("value");
		totalCoulmns=alrtDriver.findElement(By.xpath("//select[starts-with(@id,'W165')]/option[@selected='selected']")).getText();
		String Index_Page_Template_Label=null;
		
		for(int i=1;i<=Integer.parseInt(totalCoulmns);i++)
		{
			String headerValue=alrtDriver.findElement(By.xpath("//td//label[contains(.,'Label for Column "+i+"')]//following::div[1]/input")).getAttribute("value");
			Index_Page_Template_Label=Index_Page_Template_Label+","+headerValue;
		}
		if(Index_Page_Template_Label.startsWith("null"))
		{
			Index_Page_Template_Label=Index_Page_Template_Label.substring(5);
		}
		
		System.out.println(Index_Page_Template_Label);
		String Index_Page_Template="Table_"+Integer.parseInt(totalCoulmns)+"_columns";
		dataToFetch.put("Title",tableName );
		dataToFetch.put("Index_Page_Template",Index_Page_Template );
		dataToFetch.put("Index_Page_Template_Label",Index_Page_Template_Label );
		
		alrtDriver.findElement(By.id("save_and_close")).click();
		//closeContent.click();
		
		return dataToFetch;
		
	}


	private static String checkContentType(String indexchildName) {
		String fetchSiteArea;
		
		try {
			
			
			String fetchAttribute=alrtDriver.findElement(By.xpath("//a[.='"+indexchildName+"' and contains(@title,'View children')]")).getAttribute("id");
			
			String[] checkboxNumber=fetchAttribute.split("_");
			String checkboxToClick =  checkboxNumber[checkboxNumber.length-1].trim();
			
			alrtDriver.findElement(By.xpath("//input[@value='"+checkboxToClick+"']")).click();
			
			//editContent.click();
			readContent.click();
			fetchSiteArea=siteArea.getText();
			
			//alrtDriver.findElement(By.id("save_and_close")).click();
			closeContent.click();
			
			return fetchSiteArea;
		}
		catch(Exception e)
		{
			
			System.out.println("Error while determining contenet type of:"+indexchildName);
			
		}
		return null;
		
	}

	
	
	public static void fetchContentTillGrandChild(List<String> SAT_Index_Pages,String departmentName, String subDeptsUnderDeptName) throws Throwable
	{
		HashMap<String, String> wcmKeyValue1= new HashMap<String, String>();
		try
		{
			System.out.println("Fetching index pages content under Sub department "+subDeptsUnderDeptName);
			
			for(int ip=0;ip<numberOfContentsToFetch(SAT_Index_Pages);ip++)
			{
			System.out.println("SAT_Index page number"+(++ip)+" : "+SAT_Index_Pages.get(ip));
			WebElement indexPage=alrtDriver.findElement(By.xpath("//a[.='"+SAT_Index_Pages.get(ip)+"' and starts-with(@title,'View children')]"));
						
			String indexPageTitle= indexPage.getText();
					
			indexPage.click(); // FIRST INDEX PAGE :---Business continuation /Optimization clicked
			//now checking for the child content for the Index page(TCFA index page clicked)
				
			List<WebElement> allChildImages= allChildren;					
			
			//adding contents under Index pages under different lists with Child and no Child	
				//creating list for link portlets and grandchilds under Child					
			List<String> ChildHasGrandChild = new ArrayList<String>();
			List<String> IndexPageLinkPortlets = new ArrayList<String>();
									
			List<String> IsChild_Tables=new ArrayList<String>();
			List<String> IsChild_Index_pages=new ArrayList<String>();;
			List<String> IsChild_Categories=new ArrayList<String>();
			
								
				for(int cgc=1;cgc<=allChildImages.size();cgc++)
				{
					String childImageTitle=alrtDriver.findElement(By.xpath("//tr["+cgc+"]//td[2]//img[2]")).getAttribute("title");
					if(childImageTitle.contains("View children"))
						{
						String childwithGrandChild=alrtDriver.findElement(By.xpath("//tr["+cgc+"]//td[2]//img[2]/following::td[1]//a")).getText();
						ChildHasGrandChild.add(childwithGrandChild);
						}
					else 
						{
						String ChildWithNoGranChild=alrtDriver.findElement(By.xpath("//tr["+cgc+"]//td[2]//img[2]/following::td[1]//a")).getText();
						IndexPageLinkPortlets.add(ChildWithNoGranChild);
						}
				}
							
				
				System.out.println("Total child under Index page: "+indexPageTitle+" apart from Link Portlets are::"+ChildHasGrandChild.size());
			
				///writing index page(TCFA_Index_Page) Link Portlets contents
									
				for(int cng=0;cng<numberOfContentsToFetch(IndexPageLinkPortlets);cng++)
					{
						 System.out.println("fetching link portlet "+IndexPageLinkPortlets.get(cng)+" for index page::"+indexPageTitle);
						 String indexPageLinks=IndexPageLinkPortlets.get(cng);
									    	
						WebElement child12=alrtDriver.findElement(By.xpath("//a[.='"+indexPageLinks+"' and not(contains(@title, 'View children'))]"));
						 child12.click();
										 	
				    	String wcmTCID=testCaseID+testcaseNumber;
									    	
				    	wcmKeyValue1.put("Test Case ID",wcmTCID);
				    	wcmKeyValue1.put("DepartmentName",departmentName);
				    	wcmKeyValue1.put("2ndLevel",subDeptsUnderDeptName);
				    	wcmKeyValue1.put("3rdLevelIndexPage",indexPageTitle);
							    	
				   	   writeWCMToExcel(wcmKeyValue1);
				   	   writeWCMHeaderContentFinalToExcel();
				   	   testcaseNumber++;
				  	   closeContent.click();
									
					}
									
	////identifying Index page's(TCFA_Index_Page) "indexPageTitle":- Child Content Type (TCFA_Child_Index_Pag), Tables or Categories
				
				System.out.println("Now checking for Index page:"+indexPageTitle+" content type apart from link portlets");
					for(int z=0;z<ChildHasGrandChild.size();z++)
						{
							WebElement childIndexPage=alrtDriver.findElement(By.xpath("//a[.='"+ChildHasGrandChild.get(z)+"' and starts-with(@title,'View children')]"));
							String childIndexPageTitle=childIndexPage.getText(); 
							
							String childType=checkContentType(childIndexPageTitle);
							if(childType.contains("Table"))
							{
								IsChild_Tables.add(childIndexPageTitle);
							}
							else if(childType.contains("Index"))
							{
								IsChild_Index_pages.add(childIndexPageTitle);
							}
							else if(childType.contains("Site"))
							{
								IsChild_Categories.add(childIndexPageTitle);
							}
							
							System.out.println("This Child of index page "+childIndexPageTitle+" is a "+childType);
						}
					
					System.out.println("Index page: "+indexPageTitle+" has "+IsChild_Tables.size()+" Tables,"+IsChild_Categories.size()+" Categories and "+IsChild_Index_pages.size()+" Child Index Pages");
					//NOW READING INDEX PAGE TABLES
					
					for(int ct=0;ct<numberOfContentsToFetch(IsChild_Tables);ct++)
					{
						
						System.out.println("Reading content for Index Page "+indexPageTitle+" Child table "+IsChild_Tables.get(ct));// Child Table
						WebElement childTable=alrtDriver.findElement(By.xpath("//a[.='"+IsChild_Tables.get(ct)+"' and starts-with(@title,'View children')]"));
									
						String childtableTitle=childTable.getText(); 
						
						String wcmTCID=testCaseID+testcaseNumber;
				    	Map<String, String> tabledata= new HashMap<String, String>();
						
						tabledata=fetchTablesContent(childtableTitle);
						
				    	wcmKeyValue1.put("Test Case ID",wcmTCID);
				    	wcmKeyValue1.put("DepartmentName",departmentName);
				    	wcmKeyValue1.put("2ndLevel",subDeptsUnderDeptName);
				    	wcmKeyValue1.put("3rdLevelIndexPage",indexPageTitle);
						wcmKeyValue1.putAll(tabledata);
						excelOutput(wcmKeyValue1);
	      		           writeWCMHeaderContentFinalToExcel();
				           testcaseNumber++;
				       
					}/// all tables for index pages read successfully
					
					
					
			//now reading index page's child category (SALES AND OPTIMIZATION(Sub Optimization))
					for(int cc=0;cc<numberOfContentsToFetch(IsChild_Categories);cc++)
					{
						//Map<String,String> categoryContent=new HashMap<String,String>();
						System.out.println("Reading content for Category "+IsChild_Categories.get(cc)+ " of Index Page::"+indexPageTitle);// SALES
						WebElement childCategory=alrtDriver.findElement(By.xpath("//a[.='"+IsChild_Categories.get(cc)+"' and starts-with(@title,'View children')]"));
									
						String childCategoryTitle=childCategory.getText(); 
						
						childCategory.click(); 
						System.out.println("category "+childCategoryTitle+" for index page::"+indexPageTitle+" is clicked ");//SALES clicked
						
						//checking for nested category
						wcmKeyValue1.put("DepartmentName",departmentName);
				    	wcmKeyValue1.put("2ndLevel",subDeptsUnderDeptName);
				    	wcmKeyValue1.put("3rdLevelIndexPage",indexPageTitle);
				    	wcmKeyValue1.put("3rdLevelIndexPageCategories",childCategoryTitle);
						
				    	checkForNestedcategories(wcmKeyValue1);
				    	
				           //alrtDriver.findElement(By.xpath("//a[contains(.,'"+childCategory+"')]")).click();
				           alrtDriver.findElement(By.xpath("//a[contains(.,'"+indexPageTitle+"')]")).click();
					}
					
					//////now checking Child Index page contents
					for(int ici=0;ici<numberOfContentsToFetch(IsChild_Index_pages);ici++)
					{
						System.out.println("Now fetching content for Child Index Page::"+IsChild_Index_pages.get(ici)+" for index page::"+indexPageTitle);
						WebElement childIndexPage=alrtDriver.findElement(By.xpath("//a[.='"+IsChild_Index_pages.get(ici)+"' and starts-with(@title,'View children')]"));
									
						String childIndexPageTitle=childIndexPage.getText(); 

				//now checking for the inside content for the CHILD INDEX PAGE (TCFA_Child_Index_Page) 
				
				childIndexPage.click();  //TCFA Child Index Page clicked
				System.out.println("Child Index page::"+childIndexPageTitle+" clicked successfully");		
				List<String> IsGrandChild_Tables=new ArrayList<String>();
				List<String> IsGrandChildIndex_Index_pages=new ArrayList<String>();
				List<String> IsGrandChildIndex_Categories=new ArrayList<String>();
						
				List<WebElement> allChildForChildIndexPage=allChildren;
								
				List<String> childIndexPageLinkPortlet=new ArrayList<String>();
				List<String> grandChildContentForChildIndexPage=new ArrayList<String>();
								
						for(int cipc=1;cipc<=allChildForChildIndexPage.size();cipc++)
						{						
							String childIndexPageContentTitle=alrtDriver.findElement(By.xpath("//tr["+cipc+"]//td[2]//img[2]")).getAttribute("title");
							if(childIndexPageContentTitle.contains("View children"))
								{
								String childIndexPageWithGrandChild=alrtDriver.findElement(By.xpath("//tr["+cipc+"]//td[2]//img[2]/following::td[1]//a")).getText();
								grandChildContentForChildIndexPage.add(childIndexPageWithGrandChild);
								}
												
							else 
								{
								String ChildindexpageLinkPortletTitle=alrtDriver.findElement(By.xpath("//tr["+cipc+"]//td[2]//img[2]/following::td[1]//a")).getText();
								childIndexPageLinkPortlet.add(ChildindexpageLinkPortletTitle);
								}
						}
						
						
						//Now reading link portlets for CHild index page
					for(int cilp=0;cilp<numberOfContentsToFetch(childIndexPageLinkPortlet);cilp++)
						{
						 System.out.println("fetching child index page link portlet:;"+childIndexPageLinkPortlet.get(cilp));
								 String childIndexLinkPortlet=childIndexPageLinkPortlet.get(cilp);
								 WebElement childIndexLink=alrtDriver.findElement(By.xpath("//a[.='"+childIndexLinkPortlet+"' and not(contains(@title, 'View children'))]"));
								 childIndexLink.click();
								 
									String wcmTCID=testCaseID+testcaseNumber;
							    	
							    	wcmKeyValue1.put("Test Case ID",wcmTCID);
							    	wcmKeyValue1.put("DepartmentName",departmentName);
							    	wcmKeyValue1.put("2ndLevel",subDeptsUnderDeptName);
							    	wcmKeyValue1.put("3rdLevelIndexPage",indexPageTitle);
							    	wcmKeyValue1.put("3rdLevelChildIndexPage",childIndexPageTitle);
							    	   writeWCMToExcel(wcmKeyValue1);
							    	   writeWCMHeaderContentFinalToExcel();
							    	   testcaseNumber++;
							    	   closeContent.click();
								
							}
						
						
						// creating list for Child index page content apart from link portlets
						for(int gcct=0;gcct<grandChildContentForChildIndexPage.size();gcct++)
						{
							
							String gccfci=grandChildContentForChildIndexPage.get(gcct);
							String grandChildType=checkContentType(gccfci);
							if(grandChildType.contains("Table"))
							{
								IsGrandChild_Tables.add(gccfci);
							}
							
							else if(grandChildType.contains("Index"))
							{
								IsGrandChildIndex_Index_pages.add(gccfci);
							}
							else if(grandChildType.contains("Site"))
							{
								IsGrandChildIndex_Categories.add(gccfci);
							}
							
							System.out.println("This child:"+gccfci +" is a "+grandChildType);
							
						}
						
						 System.out.println("Child Index page::"+childIndexPageTitle+" has total::"+IsGrandChild_Tables.size()+" tables, "+IsGrandChildIndex_Index_pages.size()+" Index pages and "+IsGrandChildIndex_Categories.size()+" categories" );
							 
						//fetching content for Child index page's tables
						for(int gct=0;gct<numberOfContentsToFetch(IsGrandChild_Tables);gct++)
						{
							
							System.out.println("Fetching Grand Child Table:: "+IsGrandChild_Tables.get(gct)+" Content");// Child Table
							WebElement grandChildTable=alrtDriver.findElement(By.xpath("//a[.='"+IsGrandChild_Tables.get(gct)+"' and starts-with(@title,'View children')]"));
										
							String grandChildtableTitle=grandChildTable.getText(); 
							
							String wcmTCID=testCaseID+testcaseNumber;
					    	Map<String, String> tabledata= new HashMap<String, String>();
							
							tabledata=fetchTablesContent(grandChildtableTitle);
							
					    	wcmKeyValue1.put("Test Case ID",wcmTCID);
					    	wcmKeyValue1.put("DepartmentName",departmentName);
					    	wcmKeyValue1.put("2ndLevel",subDeptsUnderDeptName);
					    	wcmKeyValue1.put("3rdLevelIndexPage",indexPageTitle);
					    	wcmKeyValue1.put("3rdLevelChildIndexPage",childIndexPageTitle);
							wcmKeyValue1.putAll(tabledata);
						
							excelOutput(wcmKeyValue1);
		      		   	   writeWCMHeaderContentFinalToExcel();
					   	   testcaseNumber++;
					   	
						}
						
						///Fetching content for Child index page's categories
						
						
						for(int cc=0;cc<numberOfContentsToFetch(IsGrandChildIndex_Categories);cc++)
						{
							
							//Map<String,String> childCategoryContent=new HashMap<String,String>();
							System.out.println("Reading content for Category "+IsGrandChildIndex_Categories.get(cc)+ " of Child Index Page:"+childIndexPageTitle);// SALES
							
							WebElement grandchildCategory=alrtDriver.findElement(By.xpath("//a[.='"+IsGrandChildIndex_Categories.get(cc)+"' and starts-with(@title,'View children')]"));
										
							String childCategoryTitle=grandchildCategory.getText(); 
							
							grandchildCategory.click(); //tools and documents
							System.out.println("Category ::"+childCategoryTitle+" for Child index page:"+childIndexPageTitle+" is clicked");//Child index page first category clicked
						
							//checking for nested category
							wcmKeyValue1.put("DepartmentName",departmentName);
					    	wcmKeyValue1.put("2ndLevel",subDeptsUnderDeptName);
					    	wcmKeyValue1.put("3rdLevelIndexPage",indexPageTitle);
					    	wcmKeyValue1.put("3rdLevelChildIndexPage",childIndexPageTitle);
					    	wcmKeyValue1.put("3rdLevelChildIndexPageCategories",childCategoryTitle);
							//System.out.println(wcmKeyValue1);
					    	checkForNestedcategories(wcmKeyValue1);
					    	
					    	alrtDriver.findElement(By.xpath("//a[contains(.,'"+childIndexPageTitle+"')]")).click();
							
							
							}
								
						
					/////fetching content for GrandChild Index page content
						for(int gcip=0;gcip<numberOfContentsToFetch(IsGrandChildIndex_Index_pages);gcip++)
						{
							System.out.println("Now fetching content for Grand child Index Page::"+IsGrandChildIndex_Index_pages.get(gcip)+" under Child index page::"+childIndexPageTitle);
							WebElement grandchildIndexPage=alrtDriver.findElement(By.xpath("//a[.='"+IsGrandChildIndex_Index_pages.get(ici)+"' and starts-with(@title,'View children')]"));
												
							String grandChildIndexPageTitle=grandchildIndexPage.getText(); 

							//now checking for the inside content for the CHILD INDEX PAGE (TCFA_Child_Index_Page) 
							grandchildIndexPage.click();  //TCFA Child Index Page clicked
									
							List<WebElement> allChildFoGrandChildIndexPage=allChildren;
											
							List<String> grandChildIndexPageLinkPortlet=new ArrayList<String>();
							List<String> grandChildindexPageContent=new ArrayList<String>();
											
									for(int gcipc=1;gcipc<=allChildFoGrandChildIndexPage.size();gcipc++)
									{						
										String grandChildIndexPageContentTitle=alrtDriver.findElement(By.xpath("//tr["+gcipc+"]//td[2]//img[2]")).getAttribute("title");
										if(grandChildIndexPageContentTitle.contains("View children"))
											{
											String grandChildIndexPageWithGrandChild=alrtDriver.findElement(By.xpath("//tr["+gcipc+"]//td[2]//img[2]/following::td[1]//a")).getText();
											grandChildindexPageContent.add(grandChildIndexPageWithGrandChild);
											}
										else 
											{
											String grandChildindexpageLinkPortletTitle=alrtDriver.findElement(By.xpath("//tr["+gcipc+"]//td[2]//img[2]/following::td[1]//a")).getText();
											grandChildIndexPageLinkPortlet.add(grandChildindexpageLinkPortletTitle);
											}
									}
							
							
									////now fetching grand child index page link portlets
									for(int gclp=0;gclp<numberOfContentsToFetch(grandChildIndexPageLinkPortlet);gclp++)
									{
									 System.out.println("Fetching Grand child index page Link portlets::"+grandChildIndexPageLinkPortlet.get(gclp));
											 String grandchildContents=grandChildIndexPageLinkPortlet.get(gclp);
										    	
											 WebElement grandChildLinks=alrtDriver.findElement(By.xpath("//a[.='"+grandchildContents+"' and not(contains(@title, 'View children'))]"));
											 
											 grandChildLinks.click();
											 
											 String wcmTCID=testCaseID+testcaseNumber;
										    	
										    	wcmKeyValue1.put("Test Case ID",wcmTCID);
										    	wcmKeyValue1.put("DepartmentName",departmentName);
										    	wcmKeyValue1.put("2ndLevel",subDeptsUnderDeptName);
										    	wcmKeyValue1.put("3rdLevelIndexPage",indexPageTitle);
										    	wcmKeyValue1.put("3rdLevelChildIndexPage",childIndexPageTitle);
										    	wcmKeyValue1.put("3rdLevelGrandChildIndexPage",grandchildContents);
										    	
										    	   writeWCMToExcel(wcmKeyValue1);
										    	   writeWCMHeaderContentFinalToExcel();
										    	   testcaseNumber++;
										           closeContent.click();
										}
								 
									List<String> IsFinalChild_Tables=new ArrayList<String>();
									List<String> IsFinalChild_Categories=new ArrayList<String>();
																	
									for(int gchc=0;gchc<grandChildindexPageContent.size();gchc++)
									{
										 String grandchildWithContents=grandChildindexPageContent.get(gchc);
									    
										 WebElement finalChild=alrtDriver.findElement(By.xpath("//a[.='"+grandchildWithContents+"' and (contains(@title, 'View children'))]"));
										 
										 String grandChild=finalChild.getText();
										 String finalChildType=checkContentType(grandChild);
											if(finalChildType.contains("Table"))
											{
												IsFinalChild_Tables.add(finalChildType);
											}
											else if(finalChildType.contains("Site"))
											{
												IsFinalChild_Categories.add(finalChildType);
											}
							
									}
						
						//Now fetching table content for grand child index page
									
									for(int gctc=0;gctc<numberOfContentsToFetch(IsFinalChild_Tables);gctc++)
									{
										
										System.out.println("Fetching Grand Child Table:: "+IsFinalChild_Tables.get(gctc)+" Content");// Child Table
										WebElement finalGrandChildTable=alrtDriver.findElement(By.xpath("//a[.='"+IsFinalChild_Tables.get(gctc)+"' and starts-with(@title,'View children')]"));
													
										String finalGrandChildtableTitle=finalGrandChildTable.getText(); 
										
										String wcmTCID=testCaseID+testcaseNumber;
								    	Map<String, String> grandChildtabledata= new HashMap<String, String>();
										
								    	grandChildtabledata=fetchTablesContent(finalGrandChildtableTitle);
										
								    	wcmKeyValue1.put("Test Case ID",wcmTCID);
								    	wcmKeyValue1.put("DepartmentName",departmentName);
								    	wcmKeyValue1.put("2ndLevel",subDeptsUnderDeptName);
								    	wcmKeyValue1.put("3rdLevelIndexPage",indexPageTitle);
								    	wcmKeyValue1.put("3rdLevelChildIndexPage",childIndexPageTitle);
								    	wcmKeyValue1.put("3rdLevelGrandChildIndexPageCategories",grandChildIndexPageTitle);
								    	wcmKeyValue1.putAll(grandChildtabledata);
									
										excelOutput(wcmKeyValue1);
					      		   	   writeWCMHeaderContentFinalToExcel();
								   	   testcaseNumber++;
								   	
									}
									
									// fetching grand child categoryies content
									for(int gcfcc=0;gcfcc<numberOfContentsToFetch(IsFinalChild_Categories);gcfcc++)
									{
										System.out.println("Reading content for Category "+IsFinalChild_Categories.get(gcfcc)+ " of Grand Child Index Page"+grandChildIndexPageTitle);// SALES
										WebElement finalGrandchildCategory=alrtDriver.findElement(By.xpath("//a[.='"+IsFinalChild_Categories.get(gcfcc)+"' and starts-with(@title,'View children')]"));
													
										String grandChildCategoryTitle=finalGrandchildCategory.getText(); 
										finalGrandchildCategory.click();  //Child index page first category clicked

										wcmKeyValue1.put("DepartmentName",departmentName);
								    	wcmKeyValue1.put("2ndLevel",subDeptsUnderDeptName);
								    	wcmKeyValue1.put("3rdLevelIndexPage",indexPageTitle);
								    	wcmKeyValue1.put("3rdLevelChildIndexPage",childIndexPageTitle);
								    	wcmKeyValue1.put("3rdLevelGrandChildIndexPage",grandChildIndexPageTitle);
								    	wcmKeyValue1.put("3rdLevelGrandChildIndexPageCategories",grandChildCategoryTitle);
										
								    	checkForNestedcategories(wcmKeyValue1);
								    	alrtDriver.findElement(By.xpath("//a[contains(.,'"+grandChildIndexPageTitle+"')]")).click();
										
										
										}// end of checking for nested category or not
										
								
								 alrtDriver.findElement(By.xpath("//a[.='"+childIndexPageTitle+"']")).click();
											
								}///END OF GRANDCHILD INDEX PAGES
										
										alrtDriver.findElement(By.xpath("//a[.='"+indexPageTitle+"']")).click();// navigating back to index page Business continuation
										
									}//END OF FOR LOOP FOR ALL CHILD INDEX PAGES
						
						alrtDriver.findElement(By.xpath("//a[.='"+subDeptsUnderDeptName+"']")).click();
						
					}  //END OF FOR LOOP FOR ALL INDEX PAGES
			
			
			
			
		}//END OF TRY BLOCK
		catch(Exception e)
		{
			
			System.out.println("Error while fetching content for "+SAT_Index_Pages+" :: "+e.getMessage().toString());
		}
			
	}   //////END of fetchCOntentTillGrandChild method
	
	
	private static void checkForNestedcategories(Map<String,String> wcmKeyValue) throws Throwable{
		
		Map<String,String> wcmKeyValue1=new HashMap<String,String>();
		try
		{	
			System.out.println("now checking content for category::"+wcmKeyValue.get("3rdLevelChildIndexPageCategories"));
			
			System.out.println("category data is:"+wcmKeyValue);
			List<WebElement> categoriesContent=allChildren;
			
			List<String> categoryContent= new ArrayList<String>();
			List<String> nestedcategoryContent=new ArrayList<String>();;
 			
			for(int nc=1;nc<=categoriesContent.size();nc++)
			{
				String isNestedCategoryPresent=alrtDriver.findElement(By.xpath("//tr["+nc+"]//td[2]//img[2]")).getAttribute("title");
				String checkNestedategory=alrtDriver.findElement(By.xpath("//tr["+nc+"]//td[2]//img[2]/following::td[1]//a")).getText();
				if(isNestedCategoryPresent.contains("View children"))
				{
					nestedcategoryContent.add(checkNestedategory);
				}
				else
				{
					categoryContent.add(checkNestedategory);
				}
			}
			
			System.out.println("For category :"+wcmKeyValue.get("3rdLevelChildIndexPageCategories") +" Nested categories are: "+nestedcategoryContent.size()+" And normal content's are:"+categoryContent.size());
			
			
			if(nestedcategoryContent.size()>0)
			{
			for(int ncc=0;ncc<=numberOfContentsToFetch(nestedcategoryContent);ncc++)
			{
				System.out.println("fetching content for nested category:"+nestedcategoryContent.get(ncc));
				String nestedcategoryTitle=alrtDriver.findElement(By.xpath("//a[.='"+nestedcategoryContent.get(ncc)+"' and (contains(@title, 'View children'))]")).getText();
				//SUB SALES
				
				alrtDriver.findElement(By.xpath("//a[.='"+nestedcategoryContent.get(ncc)+"' and (contains(@title, 'View children'))]")).click();
				alrtDriver.findElement(By.xpath("//tr["+ncc+"]//td[2]//img[2]/following::td[1]//a")).click();
				
					String wcmTCID=testCaseID+testcaseNumber;
					
					wcmKeyValue1.put("Test Case ID",wcmTCID);
					wcmKeyValue1.putAll(wcmKeyValue);
		    		wcmKeyValue1.put("3rdLevelIndexPageNestedCategories",nestedcategoryTitle);
		    		writeWCMToExcel(wcmKeyValue1);
		      		        
		    	   writeWCMHeaderContentFinalToExcel();
		    	   testcaseNumber++;
		    	   closeContent.click();
		    	   
		    	   alrtDriver.findElement(By.xpath("//li[.='"+nestedcategoryTitle+"']/preceding::a[.='"+wcmKeyValue.get("3rdLevelIndexPageCategories")+"'][1]")).click();
		    	   
				}
		}
			
			for(int cc=0;cc<numberOfContentsToFetch(categoryContent);cc++)
			{
				System.out.println("fetching content for normal category content:"+categoryContent.get(cc));
				alrtDriver.findElement(By.xpath("//a[.='"+categoryContent.get(cc)+"' and not(contains(@title, 'View children'))]")).click();
			
				String wcmTCID=testCaseID+testcaseNumber;
				wcmKeyValue1.put("Test Case ID",wcmTCID);	
				wcmKeyValue1.putAll(wcmKeyValue);
					writeWCMToExcel(wcmKeyValue1);
		      		writeWCMHeaderContentFinalToExcel();
		    	   testcaseNumber++;
		    	   closeContent.click();
				}
		}
		
	catch(Exception e)
		{
		
		System.out.println("Error while checking for Nested Categories "+e.getMessage().toString());
		}
		
	}


	public static String fetchCountriesList(String country) throws Throwable
	{
		
		String countryTitle=null;
		String countryIs=null;
		try
		{
			String[] countriesArray = country.split(","); 
			 if(countriesArray.length>=2)
			 {
			
	        	for(int c=0;c<=countriesArray.length-1;c++)
				{
				    int i=countriesArray[c].indexOf("Region");
					int j=countriesArray[c].length();
					String cot=countriesArray[c].substring(i,j);
					countryTitle=countryTitle+","+cot.trim();
			
				}
	        	
	        	countryTitle=countryTitle.substring(5);
	        	return countryTitle;
			 }
			 else
			 {
				 	int i=country.indexOf("Region");
					int j=country.length();
					String cot=country.substring(i,j);
				 
					countryIs=cot;
					return countryIs;
			 }
		
		}
			catch(Exception e)
			{
			System.out.println("Error while fetching countries type from content"+e.getMessage().toString());
			
			}
		return countryTitle;
		
	}
	
	public static int numberOfContentsToFetch(List<String> listOfContent) throws Throwable{
		
		int availablecount = 0;
		try
		{
			System.out.println("Total number of contents to be fetched is::"+totalCount+" ,And content actually present is::"+listOfContent.size());
			if(listOfContent.size()<=totalCount)
			{
				availablecount=listOfContent.size();
			}
			else
			{
				availablecount=totalCount;
				
			}
			
			return availablecount;
			
		}
		catch(Exception e)
		{
			
			System.out.println("Error while comparing the number of contents to fetch "+e.getMessage().toString());
		}
		return availablecount;
		
		
		
		
	}
	
}
