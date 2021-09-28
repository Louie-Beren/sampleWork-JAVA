package main;

import utils.JS;
import utils.Read_File;
import utils.Write_To_Excel;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Set;
import java.util.concurrent.TimeUnit;


import org.openqa.selenium.By;
import org.openqa.selenium.InvalidElementStateException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class RB_FR {
	
	public static final File CONFIG_IE_FILE = new File("C:\\Automate_v1\\selenium\\IEDriverServer.exe");
	public static final File CREDENTIALS_FILE = new File("C:\\Automate_v1\\config\\credentials.txt");
	public static final File TESTDATA_FILE = new File("C:\\Automate_v1\\config\\testData.txt");
	
	public static String appURL = " "; //this part is deleted because of confidentiality
	public Boolean isAdmin = false;
	
	private InternetExplorerDriver DRIVER;
	private JavascriptExecutor EXECUTOR;
	
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
	    
		/*
		 *  VARIABLE DECLARATION
		 */
		Date startRun = new Date( );
	    SimpleDateFormat ft = new SimpleDateFormat ("yyyy-MM-dd'_'kkmmss'_'z");
	    
		String reportFileName;
		
	    RB_FR RB = new RB_FR();
	    Read_File read = new Read_File(TESTDATA_FILE);
	    
	    reportFileName = read.skipGetValueFromFile()+"_"+ft.format(startRun)+".xls";
	    
	    //	    
		RB.setDriver();
		RB.setDriverExecutor();
		InternetExplorerDriver driver = RB.getDriver();
		
		read.closeFile();		
		
	    System.out.println("Start Date/Time: " + ft.format(startRun));

		 
		driver.get(appURL);
		
		driver.manage().window().maximize();

		try
		{
			if(driver.findElement(By.tagName("html")).isDisplayed())
			{
				driver.findElement(By.tagName("html")).sendKeys(Keys.chord(Keys.CONTROL,"0"));
			}
		}
		catch(InvalidElementStateException e)
		{
			(new WebDriverWait(driver,20)).until(ExpectedConditions.presenceOfElementLocated(By.tagName("html"))).sendKeys(Keys.chord(Keys.CONTROL,"0"));
		}
		catch(StaleElementReferenceException e)
		{
			(new WebDriverWait(driver,40)).until(ExpectedConditions.presenceOfElementLocated(By.tagName("html"))).sendKeys(Keys.chord(Keys.CONTROL,"0"));
		}
		
		/* ==============================================
		 * START TESTING
		 ============================================== */
		if(RB.isRBLogin())
		{
			if(RB.isRBClickFinder())
			{
				Write_To_Excel write = new Write_To_Excel("C:\\Automate_v1\\extract\\"+reportFileName, RB.getCampaignName());
				RB.getStepProperties(write);
				write.setAutoColWidth();
				System.out.println("=========================");
				System.out.println(reportFileName+" has been generated");
			}
			
		}
		/* ==============================================
		 * END TESTING
		 ============================================== */

		Date endRun = new Date( );
		System.out.println("End Date/Time: " + ft.format(endRun));

	}
	

	private boolean isRBLogin()
	{
		InternetExplorerDriver driver = this.getDriver();
		JavascriptExecutor executor = this.getDriverExecutor();
		String thisWindow= driver.getWindowHandle();
		String newWindow;
		String userName;
		String passWd;
		String MP;
		Boolean loginResult = false;
		
		Read_File read = new Read_File(CREDENTIALS_FILE);
		Read_File read1 = new Read_File(TESTDATA_FILE);
		
		userName = read.getValueFromFile();
		passWd = read.getValueFromFile();
		MP = read1.getValueFromFile();
		
		//!! if there are no credentials
		if(userName.isEmpty() && passWd.isEmpty())
		{
			System.out.println("ERROR: Credentials.txt might be incomplete");
			return loginResult;
		}
		
		if(MP.isEmpty())
		{
			this.isAdmin = true;
		}
		read.closeFile();
		read1.closeFile();
		
		/* ==============================================
		 * LOGIN PAGE
		 ============================================== */
		
		//username
		try
		{
			WebElement LOGIN_userName_INPUT = (new WebDriverWait(driver,10)).until(ExpectedConditions.presenceOfElementLocated(By.id("username")));
			LOGIN_userName_INPUT.sendKeys(userName);
		}
		catch(Exception e)
		{
			System.out.println("ERROR: Credentials might be incorrect");
			return loginResult;
		}
		
		 
		 //password
		 WebElement LOGIN_passwrd_INPUT = driver.findElement(By.id("password"));
		 LOGIN_passwrd_INPUT.sendKeys(passWd);
		 LOGIN_passwrd_INPUT.submit();
		 
		 /*
		  * SECONDARY PAGE
		  */
		 //MP#
		 
		 WebElement LOGIN_mPrgrmNum_INPUT = (new WebDriverWait(driver,20)).until(ExpectedConditions.presenceOfElementLocated(By.id("login_data")));
		 LOGIN_mPrgrmNum_INPUT.sendKeys(MP); // 212
		 
		 WebElement LOGIN_submit_BTN = driver.findElement(By.id("loginForm"));
		 executor.executeScript("arguments[0].submit()", LOGIN_submit_BTN);
		 
		 try
		 {
			 newWindow = new WebDriverWait(driver, 20).until(new ExpectedCondition<String>(){
			   @Override
			   public String apply(WebDriver d) {
			     Set<String> handles = d.getWindowHandles();
			     handles.remove(thisWindow);
			     return handles.size() > 0 ? handles.iterator().next() : null;
			   }});
		 
			 driver.switchTo().window(newWindow);
		 }
		 catch(NoSuchElementException TOExc)
		 {
			 newWindow = new WebDriverWait(driver, 30).until(new ExpectedCondition<String>(){
				   @Override
				   public String apply(WebDriver d) {
				     Set<String> handles = d.getWindowHandles();
				     handles.remove(thisWindow);
				     return handles.size() > 0 ? handles.iterator().next() : null;
				   }});
			 
				 driver.switchTo().window(newWindow);
		 }
		 catch(Exception e)
		 {
			 System.out.println(e.getMessage());
		 }
		 
		 return true;
	}
	
	private boolean isRBClickFinder()
	{
		InternetExplorerDriver driver = this.getDriver();
		JavascriptExecutor executor = this.getDriverExecutor();
		String campaignName;
		Boolean finderResult = false;
		
		JS util = new JS();
		Read_File read1 = new Read_File(TESTDATA_FILE);
		
		campaignName = read1.skipGetValueFromFile();
		read1.closeFile();
		
		try
		{
			util.checkPageIsReady(driver,15); //20
			
			WebElement HOME_finder_BTN = (new WebDriverWait(driver,60)).until(ExpectedConditions.elementToBeClickable(By.xpath("//DIV[@id='btnFinder']/B/BUTTON")));	
			executor.executeScript("arguments[0].click();", HOME_finder_BTN);
		}
		catch(NoSuchElementException e)
		{
			util.checkPageIsReady(driver,25); //20
			
			WebElement HOME_finder_BTN = (new WebDriverWait(driver,60)).until(ExpectedConditions.elementToBeClickable(By.xpath("//DIV[@id='btnFinder']/B/BUTTON")));	
			executor.executeScript("arguments[0].click();", HOME_finder_BTN);
		}
		catch(Exception e)
		{
			System.out.println("ERROR: "+e.getMessage());
			return finderResult;
		}
		
		try
		{
			WebElement FINDER_ComponentType_INPUT = driver.findElement(By.id("finderComponentType-input"));
			executor.executeScript("arguments[0].focus()", FINDER_ComponentType_INPUT);
			util.keyPress(driver,FINDER_ComponentType_INPUT,"keydown",40,1);
		}
		catch(NoSuchElementException e)
		{
			WebElement HOME_finder_BTN = (new WebDriverWait(driver,60)).until(ExpectedConditions.elementToBeClickable(By.xpath("//DIV[@id='btnFinder']/B/BUTTON")));	
			executor.executeScript("arguments[0].click();", HOME_finder_BTN);
			
			WebElement FINDER_ComponentType_INPUT = driver.findElement(By.id("finderComponentType-input"));
			executor.executeScript("arguments[0].focus()", FINDER_ComponentType_INPUT);
			util.keyPress(driver,FINDER_ComponentType_INPUT,"keydown",40,1);
		}
		catch(Exception e)
		{
			System.out.println("ERROR: "+e.getMessage());
		}
		

		 
		 WebElement FINDER_ComponentType_LIST = driver.findElement(By.id("finderComponentType-list"));
	
		 if(this.isAdmin)
		 {
			 util.keyPress(driver,FINDER_ComponentType_LIST,"keydown",40,6);
		 }
		 else
		 {
			 util.keyPress(driver,FINDER_ComponentType_LIST,"keydown",40,3);
		 }
		 util.keyPress(driver,FINDER_ComponentType_LIST,"keypress",13,1);
		 util.checkPageIsReady(driver,8); //10
			
		 WebElement FINDER_item_1_field_INPUT = driver.findElement(By.id("adv_search_panel_item_1_field-input"));
		 executor.executeScript("arguments[0].click()", FINDER_item_1_field_INPUT);
		 util.checkPageIsReady(driver,3);
		 
		 try
		 {
			 WebElement FINDER_item_1_field_LIST = driver.findElement(By.id("adv_search_panel_item_1_field-list")); 
			 util.keyPress(driver,FINDER_item_1_field_LIST,"keydown",38,3);
			 util.keyPress(driver,FINDER_item_1_field_LIST,"keypress",13,1);
		 }
		 catch(NoSuchElementException e)
		 {
			 executor.executeScript("arguments[0].click()", FINDER_item_1_field_INPUT);
			 util.checkPageIsReady(driver,3);
			 
			 WebElement FINDER_item_1_field_LIST = driver.findElement(By.id("adv_search_panel_item_1_field-list")); 
			 util.keyPress(driver,FINDER_item_1_field_LIST,"keydown",38,3);
			 util.keyPress(driver,FINDER_item_1_field_LIST,"keypress",13,1);
		 }
		 catch(Exception e)
		 {
			 System.out.println("ERROR: "+e.getMessage());
			 return finderResult;
		 }
		 
		 
		 WebElement FINDER_item_1_operator_INPUT = driver.findElement(By.id("adv_search_panel_item_1_operator-input"));
		 executor.executeScript("arguments[0].click()", FINDER_item_1_operator_INPUT);
		 
		 WebElement FINDER_item_1_operator_LIST = driver.findElement(By.id("adv_search_panel_item_1_operator-list"));
		 util.keyPress(driver,FINDER_item_1_operator_LIST,"keydown",38,2);
		 util.keyPress(driver,FINDER_item_1_operator_LIST,"keypress",13,1);
		 
		 WebElement FINDER_CampaignName_INPUT = driver.findElement(By.id("adv_search_panel_item_1_value-input"));
		 executor.executeScript("arguments[0].value=\""+campaignName+"\";", FINDER_CampaignName_INPUT);
		 util.checkPageIsReady(driver,3);
		 
		 WebElement FINDER_Search_BTN = driver.findElement(By.xpath("//DIV[@id='adv_search_panel_search']/B/BUTTON"));
		 executor.executeScript("arguments[0].click()", FINDER_Search_BTN);
		 util.checkPageIsReady(driver,7);
		 
		 
		 try
		 {
			 WebElement FINDER_searchResult_LIST = driver.findElement(By.id("grdFinder_1"));
			 util.mousePress(driver,FINDER_searchResult_LIST,"mousedown");
		 }
		 catch(NoSuchElementException e)
		 {
			 System.out.println("WARNING: The campaign is NOT FOUND!");
			 return finderResult;
		 }
		 
		 
		 WebElement FINDER_edit_BTN = driver.findElement(By.id("btnEdit"));


		 //!! check if the edit button is enabled
		 if(FINDER_edit_BTN.getAttribute("isDisabled").equals("false"))
		 {
			 executor.executeScript("arguments[0].click()", FINDER_edit_BTN);
			 util.checkPageIsReady(driver,12);
		 }
		 else
		 {
			 System.out.println("WARNING: The campaign is LOCKED!");
			 return finderResult;
		 }
		 
		 try
		 {
			 WebElement whiteBoardCanvass = (new WebDriverWait(driver,25)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='tab_2']//DIV[@id='TRMPropertiesLayout']//DIV[@class='x-panel-bwrap']//DIV[@id='TRMWhiteboardCanvas']/DIV")));
			 executor.executeScript("arguments[0].scrollLeft = arguments[0].scrollWidth", whiteBoardCanvass);
		 }
		 catch(TimeoutException e)
		 {
			 util.checkPageIsReady(driver,20);
			 WebElement whiteBoardCanvass = driver.findElementByXPath("//DIV[@id='tab_2']//DIV[@id='TRMPropertiesLayout']//DIV[@class='x-panel-bwrap']//DIV[@id='TRMWhiteboardCanvas']/DIV");
			 executor.executeScript("arguments[0].scrollLeft = arguments[0].scrollWidth", whiteBoardCanvass);
		 }
		 catch(NoSuchElementException e)
		 {
			 util.checkPageIsReady(driver,20);
			 WebElement whiteBoardCanvass = driver.findElementByXPath("//DIV[@id='tab_2']//DIV[@id='TRMPropertiesLayout']//DIV[@class='x-panel-bwrap']//DIV[@id='TRMWhiteboardCanvas']/DIV");
			 executor.executeScript("arguments[0].scrollLeft = arguments[0].scrollWidth", whiteBoardCanvass);
		 }
			
		 return true;
	}
	
	private String getCampaignName()
	{
		InternetExplorerDriver driver = this.getDriver();
		JavascriptExecutor executor = this.getDriverExecutor();
		JS util = new JS();
		String campaignName;
		
		WebElement COMMUNICATION_properties_BTN = (new WebDriverWait(driver,35)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//TABLE[@id='btnProperties']//BUTTON")));
		util.simulateMousePress(driver,COMMUNICATION_properties_BTN,"click");
		 
		 WebElement COMMUNICATION_campaignName_INPUT = (new WebDriverWait(driver,10)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='communicationDialog']//DIV[@id='communication-tabPanel:general-tab-tab']/DIV[@id='communication-tabPanel:general-tab']/DIV[@class='x-panel-bwrap']//DIV[@class='x-form-item '][1]/DIV//INPUT[@type='text']")));
		 campaignName = COMMUNICATION_campaignName_INPUT.getAttribute("value");
		 
		 WebElement COMMUNICATION_Cancel_BTN = driver.findElementByXPath("//DIV[@id='communicationDialog']//DIV[@class='x-window-bwrap']//DIV[@class='x-window-bl']//TABLE[@id='communicationDialog_btn_cancel']//BUTTON");
		 executor.executeScript("arguments[0].click()", COMMUNICATION_Cancel_BTN);
		 
		 return campaignName;

	}
	
	private void getStepProperties(Write_To_Excel write)
	{
		InternetExplorerDriver driver = this.getDriver();
		JavascriptExecutor executor = this.getDriverExecutor();
		JS util = new JS();
		
		Integer numSTEPS = driver.findElementsByXPath("//DIV[@id='tab_2']//DIV[contains(@class,'nodeWrapper yellowNode')]").size();
		System.out.println("Number of Steps: "+numSTEPS);
		Integer currentRow = 1;
		Integer numCOLLATERAL = 0;
		Integer numOUTPUTTEMP = 0;
		//String[] properties;
		ArrayList<String> properties = new ArrayList<String>();
		
		for (int numSTEPS_Ctr=1; numSTEPS_Ctr<=numSTEPS; numSTEPS_Ctr++)
		{ 
			System.out.println("\n=====================START STEP: "+numSTEPS_Ctr+" =====================");

			if (numCOLLATERAL > numOUTPUTTEMP)
			{
				currentRow= currentRow+numCOLLATERAL;
			}
			else
			{
				currentRow= currentRow+numOUTPUTTEMP;
			}
			
			WebElement STEP_Menu_Click = (new WebDriverWait(driver,10)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='tab_2']//DIV[contains(@class,'nodeWrapper yellowNode')]["+ numSTEPS_Ctr + "]//*[@class='node_header']//*[@class='gwt-MenuBar gwt-MenuBar-horizontal']//*[@class='gwt-MenuItem gwt-MenuItem-activator']")));
			executor.executeScript("arguments[0].click()", STEP_Menu_Click);
			 
			WebElement STEP_Menu = (new WebDriverWait(driver,20)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@class='gwt-MenuBarPopup']//DIV[@class='gwt-MenuBar gwt-MenuBar-vertical']")));
			util.keyPress(driver,STEP_Menu,"keydown",40,1);
			util.keyPress(driver,STEP_Menu,"keydown",13,1);
		
			WebElement STEP_Menu_Name_INPUT = (new WebDriverWait(driver,40)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='commPlanStepDialog']//DIV[@id='TRMPropertiesLayout']//DIV[@id='step-tabPanel:general-tab']//INPUT[@name='message.message.name']")));
			System.out.println("\tStep Name: "+STEP_Menu_Name_INPUT.getAttribute("value"));
			properties.add(STEP_Menu_Name_INPUT.getAttribute("value"));
			 
			WebElement STEP_Menu_DefaultChannelClass_INPUT = driver.findElementByXPath("//DIV[@id='commPlanStepDialog']//DIV[@id='TRMPropertiesLayout']//DIV[@id='step-tabPanel:general-tab']//INPUT[@id='message-message-defaultChannelClassId-input']");
			System.out.println("\tStep Default Channel Class: "+STEP_Menu_DefaultChannelClass_INPUT.getAttribute("value"));
			properties.add(STEP_Menu_DefaultChannelClass_INPUT.getAttribute("value"));
			
			util.checkPageIsReady(driver, 2);
			WebElement STEP_Menu_DefaultChannelInstance_INPUT = driver.findElementByXPath("//DIV[@id='commPlanStepDialog']//DIV[@id='TRMPropertiesLayout']//DIV[@id='step-tabPanel:general-tab']//INPUT[@id='message-message-defaultChannelInstanceId-input']");
			System.out.println("\tStep Default Channel Instance: "+STEP_Menu_DefaultChannelInstance_INPUT.getAttribute("value"));
			properties.add(STEP_Menu_DefaultChannelInstance_INPUT.getAttribute("value"));
			
			write.writeStepProperties(currentRow,properties);
			write.closeExcelFile();
			properties.clear();
			
			/*
			 * TIMEOUT
			 */
			WebElement STEP_TIMEOUT_TAB = (new WebDriverWait(driver,20)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='commPlanStepDialog']//DIV[@id='TRMPropertiesLayout']/DIV[@id='trmPropertiesTabPanelView']//LI[@id='trmPropertiesTabPanelView__message-tabPanel:timeout-tab-tab']")));
			executor.executeScript("arguments[0].click()", STEP_TIMEOUT_TAB);
			
			this.getTimeoutProperties(currentRow,write);
			
			/*
			 * COLLATERAL 
			 */
			WebElement STEP_Collateral_TAB = (new WebDriverWait(driver,20)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='commPlanStepDialog']//DIV[@id='TRMPropertiesLayout']/DIV[@id='trmPropertiesTabPanelView']//LI[@id='trmPropertiesTabPanelView__message-tabPanel:collateral-tab-tab']")));
			executor.executeScript("arguments[0].click()", STEP_Collateral_TAB);
			
			numCOLLATERAL = this.getCollateralProperties(currentRow,write);
			
			/*
			 * OUTPUT TEMPLATE
			 */
			
			WebElement STEP_Output_TAB = (new WebDriverWait(driver,20)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='commPlanStepDialog']//DIV[@id='TRMPropertiesLayout']/DIV[@id='trmPropertiesTabPanelView']//LI[@id='trmPropertiesTabPanelView__message-tabPanel:templates-tab-tab']")));
			executor.executeScript("arguments[0].click()", STEP_Output_TAB);
			
			numOUTPUTTEMP = this.getOutputTemplateProperties(currentRow,write);
						 			 
			WebElement STEP_Cancel_BTN = (new WebDriverWait(driver,80)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='commPlanStepDialog']//DIV[@class=' x-panel-btns']//TABLE[@id='commPlanStepDialog_btn_cancel']//BUTTON")));
			executor.executeScript("arguments[0].click()", STEP_Cancel_BTN);
			System.out.println("\n=====================END STEP: "+numSTEPS_Ctr+" =====================");

		}

	}
	
	private void getTimeoutProperties(Integer row,Write_To_Excel write)
	{
		
		InternetExplorerDriver driver = this.getDriver();

		ArrayList<String> properties = new ArrayList<String>(); 

		WebElement TIMEOUT_TimeoutDays_INPUT = driver.findElementByXPath("//DIV[@id='message-tabPanel:timeout-tab-tab']//INPUT[@id='message-timeoutDelayNum-input']");
		System.out.println("\t\t\t Timeout Days: "+TIMEOUT_TimeoutDays_INPUT.getAttribute("value"));
		properties.add(TIMEOUT_TimeoutDays_INPUT.getAttribute("value"));
		
		WebElement TIMEOUT_ActionType_INPUT = driver.findElementByXPath("//DIV[@id='x-form-el-message-timeoutActionTypeCd']/DIV/INPUT");
		System.out.println("\t\t\t Action Type: "+TIMEOUT_ActionType_INPUT.getAttribute("value"));
		properties.add(TIMEOUT_ActionType_INPUT.getAttribute("value"));
		
		WebElement TIMEOUT_ActionCon_INPUT = driver.findElementByXPath("//DIV[@id='x-form-el-timeoutStepId']/DIV/INPUT");
		System.out.println("\t\t\t Action Type: "+TIMEOUT_ActionCon_INPUT.getAttribute("value"));
		properties.add(TIMEOUT_ActionCon_INPUT.getAttribute("value"));
		
		write.writeTimeoutProperties(row,properties);
		write.closeExcelFile();
		properties.clear();
	}
	
	private Integer getCollateralProperties(Integer row,Write_To_Excel write)
	{
		
		InternetExplorerDriver driver = this.getDriver();
		JavascriptExecutor executor = this.getDriverExecutor();
		JS util = new JS();
		String COLLATERAL_Email_name;
		String COLLATERAL_Email_SL_val;
		String COLLATERAL_Email_Info_Name_val;
		
		ArrayList<String> properties = new ArrayList<String>(); 
		Integer numCOLLATERAL = driver.findElementsByXPath("//DIV[@id='TRMPropertiesLayout']//DIV[@class='x-tab-panel-body x-tab-panel-body-top']/DIV[@id='message-tabPanel:collateral-tab-tab']//DIV[@id='commPlanStepCollateralGridPanel']//DIV[@id='commPlanStepCollateralGridPanel-grid']//DIV[@class='x-grid3-scroller']/DIV[@class='x-grid3-body']//DIV[contains(@id,'commPlanStepCollateralGridPanel-grid')]").size();
		System.out.println("\t\t***Number of Collateral(s):"+numCOLLATERAL);

		 
		 for(Integer numCOLLATERAL_Ctr=1;numCOLLATERAL_Ctr<=numCOLLATERAL;numCOLLATERAL_Ctr++)
		 {
			 WebElement COLLATERAL_Row_CHKBOX = driver.findElementByXPath("//DIV[@id='TRMPropertiesLayout']//DIV[@class='x-tab-panel-body x-tab-panel-body-top']/DIV[@id='message-tabPanel:collateral-tab-tab']//DIV[@id='commPlanStepCollateralGridPanel']//DIV[@id='commPlanStepCollateralGridPanel-grid']//DIV[@class='x-grid3-scroller']/DIV[@class='x-grid3-body']//DIV[@id='commPlanStepCollateralGridPanel-grid_"+numCOLLATERAL_Ctr+"']//DIV[@class='x-grid3-row-checker']");
			 System.out.println("\t\t===================== START COLLATERAL: "+numCOLLATERAL_Ctr);
			 util.simulateMousePress(driver,COLLATERAL_Row_CHKBOX,"mousedown");
			 
			 WebElement COLLATERAL_Open_BTN = driver.findElementByXPath("//DIV[@id='TRMPropertiesLayout']//DIV[@class='x-tab-panel-body x-tab-panel-body-top']/DIV[@id='message-tabPanel:collateral-tab-tab']//DIV[@id='commPlanStepCollateralGridPanel']//DIV[@id='commPlanStepCollateralGridPanel-topToolbar']//DIV[@id='toolbarOpen']//BUTTON");
			 executor.executeScript("arguments[0].click()", COLLATERAL_Open_BTN);
			 
			 WebElement COLLATERAL_Name_INPUT = driver.findElementByXPath("//DIV[@id='commPlanMessageCollateralDialog']//DIV[@id='TRMPropertiesLayout']//DIV[@id='collateral-tabPanel:general-tab-tab']//DIV[@class='x-panel-bwrap']//INPUT[@name='name']");
			 System.out.println("\t\t\t Collateral Name: "+COLLATERAL_Name_INPUT.getAttribute("value"));
			 properties.add(COLLATERAL_Name_INPUT.getAttribute("value"));
			 
			 WebElement COLLATERAL_Type_INPUT = driver.findElementByXPath("//DIV[@id='commPlanMessageCollateralDialog']//DIV[@id='TRMPropertiesLayout']//DIV[@id='collateral-tabPanel:general-tab-tab']//DIV[@class='x-panel-bwrap']//INPUT[@id='collateralTypeCd-input']");
			 System.out.println("\t\t\t Collateral Type: "+COLLATERAL_Type_INPUT.getAttribute("value"));
			 properties.add(COLLATERAL_Type_INPUT.getAttribute("value"));
			 
			 
			 WebElement COLLATERAL_Email_BTN = driver.findElementByXPath("//DIV[@id='commPlanMessageCollateralDialog']//DIV[@id='TRMPropertiesLayout']//DIV[@id='collateral-tabPanel:general-tab-tab']//DIV[@class='x-panel-bwrap']//DIV[@id='emailId-open']//BUTTON");
			 COLLATERAL_Email_name = COLLATERAL_Email_BTN.getText();
			 executor.executeScript("arguments[0].click()", COLLATERAL_Email_BTN);
			 
			 if(COLLATERAL_Email_name.equals("< Not Defined >"))
			 {
				 properties.add("none");
				 properties.add("none");
			 }
			 else
			 {
				 try
				 {
					 WebElement COLLATERAL_Email_FRAME = (new WebDriverWait(driver,80)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='tab_2']//IFRAME")));
				 }
				 catch(NoSuchElementException e)
				 {
					 WebElement COLLATERAL_Email_FRAME = (new WebDriverWait(driver,180)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='tab_2']//IFRAME")));
				 }
				 
				 WebElement COLLATERAL_Email_Info_BTN = (new WebDriverWait(driver,20)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='headerPanelId']//DIV[@id='CommonPropertiesDialogButtonId']//BUTTON[@class='x-btn-text button']")));
				 executor.executeScript("arguments[0].click()", COLLATERAL_Email_Info_BTN);
				
				 WebElement COLLATERAL_Email_Info_Name = driver.findElementByXPath("//DIV[@id='commonPropertiesDialogId']//SPAN[@id='commonPropertiesDialogId_Heading-label']");
				 COLLATERAL_Email_Info_Name_val = COLLATERAL_Email_Info_Name.getText();
				 System.out.println("\t\t\t Subject Line: "+COLLATERAL_Email_Info_Name_val.substring(5));
				 properties.add(COLLATERAL_Email_Info_Name_val.substring(5));
				 
				 WebElement COLLATERAL_Email_Info_DMCId = driver.findElementByXPath("//DIV[@id='commonPropertiesDialogId']//DIV[@id='commonPropertiesDialogObjectIdId']");
				 System.out.println("\t\t\t DMC ID: "+ COLLATERAL_Email_Info_DMCId.getText());
				 properties.add(COLLATERAL_Email_Info_DMCId.getText());
				 
				 WebElement COLLATERAL_Email_Info_Close_BTN = driver.findElementByXPath("//TABLE[@id='commonPropertiesDialogId_btn_close']//BUTTON");
				 executor.executeScript("arguments[0].click()", COLLATERAL_Email_Info_Close_BTN);
				 
				 driver.switchTo().frame(0);
				 
				 WebElement COLLATERAL_Email_SL;
				 
				 if(this.isAdmin)
				 {
					 COLLATERAL_Email_SL = (new WebDriverWait(driver,180)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@class='option-row']//TEXTAREA[@id='message-subject']")));
					 COLLATERAL_Email_SL_val = COLLATERAL_Email_SL.getText();
					 System.out.println("\t\t\t Subject Line: "+COLLATERAL_Email_SL.getText());
					 properties.add(COLLATERAL_Email_SL.getText());
				 }
				 else
				 {
					 COLLATERAL_Email_SL = (new WebDriverWait(driver,180)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@class='preview-subject-header']/DIV[1]")));
					 COLLATERAL_Email_SL_val = COLLATERAL_Email_SL.getText();
					 System.out.println("\t\t\t Subject Line: "+COLLATERAL_Email_SL_val.substring(9));
					 properties.add(COLLATERAL_Email_SL_val.substring(9));
				 }
				 
				 driver.switchTo().defaultContent();
				
			 }
			  
			 WebElement COLLATERAL_Email_Back_BTN = (new WebDriverWait(driver,60)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//table[@id='backButtonId']//BUTTON")));
			 util.simulateMousePress(driver,COLLATERAL_Email_Back_BTN,"click");

			 /*
			  * CONTENT/INCENTIVE
			  */

			 WebElement COLLATERAL_Content_TAB = (new WebDriverWait(driver,60)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='TRMPropertiesLayout']//LI[@id='trmPropertiesTabPanelView__10007vzdlznb-tab']")));
			 executor.executeScript("arguments[0].click()", COLLATERAL_Content_TAB);
			 
			 WebElement COLLATERAL_Content_value_INPUT = driver.findElementByXPath("//DIV[@id='10007vzdlznb-tab']//INPUT[@name='ex_collateral.cntnt_inctv_id']");

			 
			 if (COLLATERAL_Content_value_INPUT.getAttribute("value").toString().equals("Edit"))
			 {
				 WebElement COLLATERAL_Content_IMG = driver.findElementByXPath("//DIV[@id='10007vzdlznb-tab']//INPUT[@name='ex_collateral.cntnt_inctv_id']/following-sibling::IMG");
				 executor.executeScript("arguments[0].click()", COLLATERAL_Content_IMG);
				 
				 util.checkPageIsReady(driver, 10);
				 
				 WebElement COLLATERAL_Content_Incentive_TEXTAREA = driver.findElementByXPath("//DIV[@id='trmDialog']//TEXTAREA[@id='SchemaSelectorDesc_TextArea_id-input']");
				 System.out.println("\t\t\t Incentive: "+COLLATERAL_Content_Incentive_TEXTAREA.getText());
				 properties.add(COLLATERAL_Content_Incentive_TEXTAREA.getText());
				 
				 WebElement COLLATERAL_Content_Incentive_BTN = driver.findElementByXPath("//TABLE[@id='trmDialog_btn_cancel']//BUTTON");
				 executor.executeScript("arguments[0].click()", COLLATERAL_Content_Incentive_BTN);
			 }
			 else
			 {
				 properties.add("NONE");
			 }
					 
			 
			 /*
			  * AB TESTING
			  */
			 WebElement COLLATERAL_ABTest_TAB = (new WebDriverWait(driver,60)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='TRMPropertiesLayout']//LI[@id='trmPropertiesTabPanelView__1000gg1v5c85-tab']")));
			 executor.executeScript("arguments[0].click()", COLLATERAL_ABTest_TAB);
			 
			 try
			 {
				 WebElement COLLATERAL_ABTest_Placeholder_INPUT = driver.findElementByXPath("//DIV[@id='1000gg1v5c85-tab']//DIV[@id='ex_collateral-test_placeholder_ind']/INPUT");
				 System.out.println("\t\t\t Placeholder: "+COLLATERAL_ABTest_Placeholder_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_Placeholder_INPUT.getAttribute("value"));
			 }
			 catch(StaleElementReferenceException e)
			 {
				 util.checkPageIsReady(driver, 3);
				 WebElement COLLATERAL_ABTest_Placeholder_INPUT = driver.findElementByXPath("//DIV[@id='1000gg1v5c85-tab']//DIV[@id='ex_collateral-test_placeholder_ind']/INPUT");
				 System.out.println("\t\t\t Placeholder: "+COLLATERAL_ABTest_Placeholder_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_Placeholder_INPUT.getAttribute("value"));
			 }

			 try
			 {
				 WebElement COLLATERAL_ABTest_RequiredResp_INPUT = (new WebDriverWait(driver,60)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='1000gg1v5c85-tab']//INPUT[@name='ex_collateral.test_response_pct']")));
				 System.out.println("\t\t\t Required Response %: "+COLLATERAL_ABTest_RequiredResp_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_RequiredResp_INPUT.getAttribute("value"));
			 }
			 catch(StaleElementReferenceException e)
			 {
				 util.checkPageIsReady(driver, 5);
				 WebElement COLLATERAL_ABTest_RequiredResp_INPUT = (new WebDriverWait(driver,60)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='1000gg1v5c85-tab']//INPUT[@name='ex_collateral.test_response_pct']")));
				 System.out.println("\t\t\t Required Response %: "+COLLATERAL_ABTest_RequiredResp_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_RequiredResp_INPUT.getAttribute("value"));
			 }
			 
			 try
			 {
				 WebElement COLLATERAL_ABTest_ResponseTyp_INPUT = (new WebDriverWait(driver,60)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='1000gg1v5c85-tab']//INPUT[@name='ex_collateral.test_response_type_cd']")));
				 System.out.println("\t\t\t Response Type: "+COLLATERAL_ABTest_ResponseTyp_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_ResponseTyp_INPUT.getAttribute("value"));
			 }
			 catch(StaleElementReferenceException e)
			 {
				 util.checkPageIsReady(driver, 5);
				 WebElement COLLATERAL_ABTest_ResponseTyp_INPUT = (new WebDriverWait(driver,60)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='1000gg1v5c85-tab']//INPUT[@name='ex_collateral.test_response_type_cd']")));
				 System.out.println("\t\t\t Response Type: "+COLLATERAL_ABTest_ResponseTyp_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_ResponseTyp_INPUT.getAttribute("value"));
			 }
			 
			 try
			 {
				 WebElement COLLATERAL_ABTest_MinWait_INPUT = driver.findElementByXPath("//DIV[@id='1000gg1v5c85-tab']//INPUT[@name='ex_collateral.test_minimum_wait_num']");
				 System.out.println("\t\t\t Min. Wait Time: "+COLLATERAL_ABTest_MinWait_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_MinWait_INPUT.getAttribute("value"));
			 }
			 catch(StaleElementReferenceException e)
			 {
				 util.checkPageIsReady(driver, 2);
				 WebElement COLLATERAL_ABTest_MinWait_INPUT = driver.findElementByXPath("//DIV[@id='1000gg1v5c85-tab']//INPUT[@name='ex_collateral.test_minimum_wait_num']");
				 System.out.println("\t\t\t Min. Wait Time: "+COLLATERAL_ABTest_MinWait_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_MinWait_INPUT.getAttribute("value")); 
			 }
			 
			 try
			 {
				 WebElement COLLATERAL_ABTest_MaxWait_INPUT = driver.findElementByXPath("//DIV[@id='1000gg1v5c85-tab']//INPUT[@name='ex_collateral.test_maximum_wait_num']");
				 System.out.println("\t\t\t Max. Wait Time: "+COLLATERAL_ABTest_MaxWait_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_MaxWait_INPUT.getAttribute("value"));
			 }
			 catch(StaleElementReferenceException e)
			 {
				 util.checkPageIsReady(driver, 2);
				 WebElement COLLATERAL_ABTest_MaxWait_INPUT = driver.findElementByXPath("//DIV[@id='1000gg1v5c85-tab']//INPUT[@name='ex_collateral.test_maximum_wait_num']");
				 System.out.println("\t\t\t Max. Wait Time: "+COLLATERAL_ABTest_MaxWait_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_MaxWait_INPUT.getAttribute("value"));
			 }
			 
			
			 try
			 {
				 WebElement COLLATERAL_ABTest_Candididate1_INPUT = driver.findElementByXPath("//DIV[@id='1000gg1v5c85-tab']//INPUT[@name='ex_collateral.candidate_1_coll_id']");
				 System.out.println("\t\t\t Candidate 1: "+COLLATERAL_ABTest_Candididate1_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_Candididate1_INPUT.getAttribute("value"));
			 }
			 catch(StaleElementReferenceException e)
			 {
				 util.checkPageIsReady(driver, 2);
				 WebElement COLLATERAL_ABTest_Candididate1_INPUT = driver.findElementByXPath("//DIV[@id='1000gg1v5c85-tab']//INPUT[@name='ex_collateral.candidate_1_coll_id']");
				 System.out.println("\t\t\t Candidate 1: "+COLLATERAL_ABTest_Candididate1_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_Candididate1_INPUT.getAttribute("value"));
			 }
			 
			 
			 
			 try
			 {
				 WebElement COLLATERAL_ABTest_Candididate2_INPUT = driver.findElementByXPath("//DIV[@id='1000gg1v5c85-tab']//INPUT[@name='ex_collateral.candidate_2_coll_id']");
				 System.out.println("\t\t\t Candidate 2: "+COLLATERAL_ABTest_Candididate2_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_Candididate2_INPUT.getAttribute("value")); 
			 }
			 catch(StaleElementReferenceException e)
			 {
				 util.checkPageIsReady(driver, 2);
				 WebElement COLLATERAL_ABTest_Candididate2_INPUT = driver.findElementByXPath("//DIV[@id='1000gg1v5c85-tab']//INPUT[@name='ex_collateral.candidate_2_coll_id']");
				 System.out.println("\t\t\t Candidate 2: "+COLLATERAL_ABTest_Candididate2_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_Candididate2_INPUT.getAttribute("value"));
			 }
			
			 
			 try
			 {
				 WebElement COLLATERAL_ABTest_Candididate3_INPUT = driver.findElementByXPath("//DIV[@id='1000gg1v5c85-tab']//INPUT[@name='ex_collateral.candidate_3_coll_id']");
				 System.out.println("\t\t\t Candidate 3: "+COLLATERAL_ABTest_Candididate3_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_Candididate3_INPUT.getAttribute("value"));
			 }
			 catch(StaleElementReferenceException e)
			 {
				 util.checkPageIsReady(driver, 2);
				 WebElement COLLATERAL_ABTest_Candididate3_INPUT = driver.findElementByXPath("//DIV[@id='1000gg1v5c85-tab']//INPUT[@name='ex_collateral.candidate_3_coll_id']");
				 System.out.println("\t\t\t Candidate 3: "+COLLATERAL_ABTest_Candididate3_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_Candididate3_INPUT.getAttribute("value"));
			 }
			
			 
			 try
			 {
				 WebElement COLLATERAL_ABTest_Candididate4_INPUT = driver.findElementByXPath("//DIV[@id='1000gg1v5c85-tab']//INPUT[@name='ex_collateral.candidate_4_coll_id']");
				 System.out.println("\t\t\t Candidate 4: "+COLLATERAL_ABTest_Candididate4_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_Candididate4_INPUT.getAttribute("value"));
			 }
			 catch(StaleElementReferenceException e)
			 {
				 util.checkPageIsReady(driver, 2);
				 WebElement COLLATERAL_ABTest_Candididate4_INPUT = driver.findElementByXPath("//DIV[@id='1000gg1v5c85-tab']//INPUT[@name='ex_collateral.candidate_4_coll_id']");
				 System.out.println("\t\t\t Candidate 4: "+COLLATERAL_ABTest_Candididate4_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_Candididate4_INPUT.getAttribute("value"));
			 }
			 
			 
			 try
			 {
				 WebElement COLLATERAL_ABTest_Candididate5_INPUT = driver.findElementByXPath("//DIV[@id='1000gg1v5c85-tab']//INPUT[@name='ex_collateral.candidate_5_coll_id']");
				 System.out.println("\t\t\t Candidate 5: "+COLLATERAL_ABTest_Candididate5_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_Candididate5_INPUT.getAttribute("value"));
			 }
			 catch(StaleElementReferenceException e)
			 {
				 util.checkPageIsReady(driver, 2);
				 WebElement COLLATERAL_ABTest_Candididate5_INPUT = driver.findElementByXPath("//DIV[@id='1000gg1v5c85-tab']//INPUT[@name='ex_collateral.candidate_5_coll_id']");
				 System.out.println("\t\t\t Candidate 5: "+COLLATERAL_ABTest_Candididate5_INPUT.getAttribute("value"));
				 properties.add(COLLATERAL_ABTest_Candididate5_INPUT.getAttribute("value"));
			 }
			 
			 
			 write.writeCollateralProperties(row+numCOLLATERAL_Ctr-1,properties);
			 write.closeExcelFile();
			 properties.clear();
			 
			 WebElement STEP_Collateral_Cancel_BTN = (new WebDriverWait(driver,60)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='commPlanMessageCollateralDialog']//TABLE[@id='commPlanMessageCollateralDialog_btn_cancel']//BUTTON")));
			 executor.executeScript("arguments[0].click()", STEP_Collateral_Cancel_BTN);
			 util.checkPageIsReady(driver,3); //20 
			 
			 System.out.println("\t\t=====================END COLLATERAL: "+numCOLLATERAL_Ctr);
		 }
		 return numCOLLATERAL;

	}
	
	private Integer getOutputTemplateProperties(Integer row,Write_To_Excel write)
	{
		
		InternetExplorerDriver driver = this.getDriver();
		JavascriptExecutor executor = this.getDriverExecutor();
		JS util = new JS();
		ArrayList<String> properties = new ArrayList<String>(); 
		Integer numOUTPUT = driver.findElementsByXPath("//DIV[@id='TRMPropertiesLayout']//DIV[@class='x-tab-panel-body x-tab-panel-body-top']/DIV[@id='message-tabPanel:templates-tab-tab']//DIV[@id='commPlanStepOutputGridPanel']//DIV[@id='commPlanStepOutputGridPanel-grid']//DIV[@class='x-grid3-scroller']/DIV[@class='x-grid3-body']//DIV[contains(@id,'commPlanStepOutputGridPanel-grid')]").size();
		System.out.println("\n\t\t***Number of Output Template(s):"+numOUTPUT);
		 
		 for(Integer numOUTPUT_Ctr=1;numOUTPUT_Ctr<=numOUTPUT;numOUTPUT_Ctr++)
		 {
			 System.out.println("\t\t=====================START OUTPUT TEMPLATE: "+numOUTPUT_Ctr);
			 
			 try
			 {
				 WebElement OUTPUT_Row_CHKBOX = (new WebDriverWait(driver,40)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='TRMPropertiesLayout']//DIV[@class='x-tab-panel-body x-tab-panel-body-top']/DIV[@id='message-tabPanel:templates-tab-tab']//DIV[@id='commPlanStepOutputGridPanel']//DIV[@id='commPlanStepOutputGridPanel-grid']//DIV[@class='x-grid3-scroller']/DIV[@class='x-grid3-body']//DIV[@id='commPlanStepOutputGridPanel-grid_"+numOUTPUT_Ctr+"']//DIV[@class='x-grid3-row-checker']")));
				 util.simulateMousePress(driver,OUTPUT_Row_CHKBOX,"mousedown");
			 }
			 catch(TimeoutException e)
			 {
				 WebElement OUTPUT_Row_CHKBOX = (new WebDriverWait(driver,40)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='TRMPropertiesLayout']//DIV[@class='x-tab-panel-body x-tab-panel-body-top']/DIV[@id='message-tabPanel:templates-tab-tab']//DIV[@id='commPlanStepOutputGridPanel']//DIV[@id='commPlanStepOutputGridPanel-grid']//DIV[@class='x-grid3-scroller']/DIV[@class='x-grid3-body']//DIV[@id='commPlanStepOutputGridPanel-grid_"+numOUTPUT_Ctr+"']//DIV[@class='x-grid3-row-checker']")));
				 util.simulateMousePress(driver,OUTPUT_Row_CHKBOX,"mousedown");
			 }
			 catch(StaleElementReferenceException e)
			 {
				 WebElement OUTPUT_Row_CHKBOX = (new WebDriverWait(driver,40)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='TRMPropertiesLayout']//DIV[@class='x-tab-panel-body x-tab-panel-body-top']/DIV[@id='message-tabPanel:templates-tab-tab']//DIV[@id='commPlanStepOutputGridPanel']//DIV[@id='commPlanStepOutputGridPanel-grid']//DIV[@class='x-grid3-scroller']/DIV[@class='x-grid3-body']//DIV[@id='commPlanStepOutputGridPanel-grid_"+numOUTPUT_Ctr+"']//DIV[@class='x-grid3-row-checker']")));
				 util.simulateMousePress(driver,OUTPUT_Row_CHKBOX,"mousedown");
			 }
			 
			 WebElement OUTPUT_Open_BTN = driver.findElementByXPath("//DIV[@id='TRMPropertiesLayout']//DIV[@class='x-tab-panel-body x-tab-panel-body-top']/DIV[@id='message-tabPanel:templates-tab-tab']//DIV[@id='commPlanStepOutputGridPanel']//DIV[@id='commPlanStepOutputGridPanel-topToolbar']//DIV[@id='toolbarOpen']//BUTTON");
			 executor.executeScript("arguments[0].click()", OUTPUT_Open_BTN);
			 
			 try
			 {
				
				 (new WebDriverWait(driver,40)).until(ExpectedConditions.frameToBeAvailableAndSwitchToIt(By.xpath("//DIV[@id='tab_2']//IFRAME"))); //"IFrameView0"

			 }
			 catch(NoSuchElementException e)
			 {
				 (new WebDriverWait(driver,80)).until(ExpectedConditions.frameToBeAvailableAndSwitchToIt(By.xpath("//DIV[@id='tab_2']//IFRAME"))); //"IFrameView0"

				 driver.switchTo().frame(0);
			 }

			 
			 try
			 {
				 WebElement OUTPUT_Name_INPUT = (new WebDriverWait(driver,20)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='communicationGeneralProperties']//INPUT[@id='trmName']")));
				 System.out.println("\t\t\t Output Name: "+OUTPUT_Name_INPUT.getAttribute("value"));
				 properties.add(OUTPUT_Name_INPUT.getAttribute("value"));
			 }
			 
			 catch(TimeoutException e)
			 {
				 WebElement OUTPUT_Name_INPUT = (new WebDriverWait(driver,40)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='communicationGeneralProperties']//INPUT[@id='trmName']")));
				 System.out.println("\t\t\t Output Name: "+OUTPUT_Name_INPUT.getAttribute("value"));
				 properties.add(OUTPUT_Name_INPUT.getAttribute("value"));
			 }
			 catch(NoSuchElementException e)
			 {
				 WebElement OUTPUT_Name_INPUT = (new WebDriverWait(driver,180)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='communicationGeneralProperties']//INPUT[@id='trmName']")));
				 System.out.println("\t\t\t Output Name: "+OUTPUT_Name_INPUT.getAttribute("value"));
				 properties.add(OUTPUT_Name_INPUT.getAttribute("value"));
			 }

			 

			 Select OUTPUT_Type_SEL = new Select(driver.findElementByXPath("//DIV[@id='communicationGeneralProperties']//SELECT[@id='outputDestControl']"));
			 util.checkPageIsReady(driver,3);
			 System.out.println("\t\t\t Output Type: "+OUTPUT_Type_SEL.getFirstSelectedOption().getText());
			 properties.add(OUTPUT_Type_SEL.getFirstSelectedOption().getText());

			 Select OUTPUT_DelSetting_SEL = new Select(driver.findElementByXPath("//DIV[@id='communicationGeneralProperties']//SELECT[@id='groupSettingsTemplateControl']"));
			 System.out.println("\t\t\t Output Delivery Setting: "+OUTPUT_DelSetting_SEL.getFirstSelectedOption().getText());
			 properties.add(OUTPUT_DelSetting_SEL.getFirstSelectedOption().getText());

			 WebElement OUTPUT_DelSetting_BTN = driver.findElementByXPath("//DIV[@id='communicationGeneralProperties']//INPUT[@id='dmcViewDeliverySetting']");
			 executor.executeScript("arguments[0].onclick()", OUTPUT_DelSetting_BTN);
			 driver.switchTo().defaultContent();
			 
			 try
			 {
				 util.checkPageIsReady(driver,3);
				 (new WebDriverWait(driver,80)).until(ExpectedConditions.frameToBeAvailableAndSwitchToIt(By.xpath("//DIV[@id='trmDialog']//IFRAME")));

			 }
			 catch(NoSuchElementException e)
			 {
				 util.checkPageIsReady(driver,7);
				 (new WebDriverWait(driver,80)).until(ExpectedConditions.frameToBeAvailableAndSwitchToIt(By.xpath("//DIV[@id='trmDialog']//IFRAME")));
			 }
			 

			 Select OUTPUT_DelSetting_DomainID_SEL = new Select((new WebDriverWait(driver,60)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='tabs-general']//SELECT[@name='domainid']"))));
			 util.checkPageIsReady(driver,3);
			 System.out.println("\t\t\t Output Domain ID: "+OUTPUT_DelSetting_DomainID_SEL.getFirstSelectedOption().getText());
			 properties.add(OUTPUT_DelSetting_DomainID_SEL.getFirstSelectedOption().getText());
			 
			 Select OUTPUT_DelSetting_Speed_SEL = new Select((new WebDriverWait(driver,40)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='tabs-general']//SELECT[@name='sendoutSpecificAmountSpeed']"))));
			 System.out.println("\t\t\t Output Speed: "+OUTPUT_DelSetting_Speed_SEL.getFirstSelectedOption().getText());
			 properties.add(OUTPUT_DelSetting_Speed_SEL.getFirstSelectedOption().getText());
			 
			 try
			 {
				 WebElement OUTPUT_DelSetting_RplyHndl_TAB = (new WebDriverWait(driver,15)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='tabs']//A[@href='#tabs-reply-handling']")));
				 util.simulateMousePress(driver,OUTPUT_DelSetting_RplyHndl_TAB,"click");
			 }
			 catch(NoSuchElementException e)
			 {
				 util.checkPageIsReady(driver,3);
				 WebElement OUTPUT_DelSetting_RplyHndl_TAB = (new WebDriverWait(driver,15)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='tabs']//A[@href='#tabs-reply-handling']")));
				 util.simulateMousePress(driver,OUTPUT_DelSetting_RplyHndl_TAB,"click");
			 }
			 

			 Select OUTPUT_DelSetting_RplyHndl_EmailReply_SEL = new Select((new WebDriverWait(driver,40)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='tabs-reply-handling']//SELECT[@name='grpmsgrpl']"))));
			 util.checkPageIsReady(driver,4);
			 System.out.println("\t\t\t Email Reply Handling: "+OUTPUT_DelSetting_RplyHndl_EmailReply_SEL.getFirstSelectedOption().getText());
			 properties.add(OUTPUT_DelSetting_RplyHndl_EmailReply_SEL.getFirstSelectedOption().getText());
			 
			 WebElement OUTPUT_DelSetting_RplyHndl_FrndName_INPUT = driver.findElementByXPath("//DIV[@id='email-defined-address']//INPUT[@name='grpmsgfromnam']");
			 System.out.println("\t\t\t Email Friendly name: "+OUTPUT_DelSetting_RplyHndl_FrndName_INPUT.getAttribute("value"));
			 properties.add(OUTPUT_DelSetting_RplyHndl_FrndName_INPUT.getAttribute("value"));
			 
			 WebElement OUTPUT_DelSetting_RplyHndl_FrmEName_INPUT = driver.findElementByXPath("//DIV[@id='email-defined-address']//INPUT[@name='grpmsgfromeml']");
			 System.out.println("\t\t\t Email From name: "+OUTPUT_DelSetting_RplyHndl_FrmEName_INPUT.getAttribute("value"));
			 properties.add(OUTPUT_DelSetting_RplyHndl_FrmEName_INPUT.getAttribute("value"));
			 
			 driver.switchTo().defaultContent();
			 
			 WebElement OUTPUT_DelSetting_Close_BTN = driver.findElementByXPath("//DIV[@id='trmDialog']//DIV[@class=' x-nodrag x-tool-close x-tool x-component']");
			 executor.executeScript("arguments[0].click()", OUTPUT_DelSetting_Close_BTN);
			 driver.switchTo().frame(0);
			 
			 WebElement OUTPUT_SendingToEmail_RADIO = driver.findElementByXPath("//TR[@id='sendToTypeSelectorRow']//INPUT[@id='sendToTypeEmail']");

			 
			 if(OUTPUT_SendingToEmail_RADIO.isSelected())
			 {
				 System.out.println("\t\t\t Sending To: Email"); 
				 properties.add("Email");
			 }
			 else
			 {
				 System.out.println("\t\t\t Sending To: SMS");
				 properties.add("SMS");
			 }
			 
			 Select OUTPUT_EmailAddress_SEL = new Select(driver.findElementByXPath("//TR[@id='sendToAttributeSelectorRow']//SELECT[@id='toAttributeControl']"));
			 util.checkPageIsReady(driver,3);
			 System.out.println("\t\t\t Email Address Field: "+OUTPUT_EmailAddress_SEL.getFirstSelectedOption().getText());
			 properties.add(OUTPUT_EmailAddress_SEL.getFirstSelectedOption().getText());
			 
			 WebElement OUTPUT_DaysAfterProc_INPUT = driver.findElementByXPath("//TR[@id='sendEmailProcessingRow']//INPUT[@id='dmcSendDaysAfterControl']");
			 System.out.println("\t\t\t # of Days Processing: "+OUTPUT_DaysAfterProc_INPUT.getAttribute("value"));
			 properties.add(OUTPUT_DaysAfterProc_INPUT.getAttribute("value"));
			 
			 WebElement OUTPUT_SendMessageTime_INPUT = driver.findElementByXPath("//TR[@id='sendEmailMessageTimeRow']//SPAN[@id='widDmcSendTime']//INPUT[@type='text']");
			 System.out.println("\t\t\t Send Message Time: "+OUTPUT_SendMessageTime_INPUT.getAttribute("value"));
			 properties.add(OUTPUT_SendMessageTime_INPUT.getAttribute("value"));
			 
			 write.writeOutputTempProperties(row+numOUTPUT_Ctr-1,properties);
			 write.closeExcelFile();
			 properties.clear();
			 
			 
			 
			 WebElement OUTPUT_Back_BTN = (new WebDriverWait(driver,5)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//DIV[@id='trmAppContainer']//DIV[@id='backButton']")));
			 util.simulateMousePress(driver,OUTPUT_Back_BTN,"click");
			 util.checkPageIsReady(driver,5); //20 
			 driver.switchTo().defaultContent();
			 System.out.println("\t\t=====================END OUTPUT TEMPLATE: "+numOUTPUT_Ctr);

		 }
		 return numOUTPUT;

	}
	public void setDriver()
	{
		DesiredCapabilities capabilities = DesiredCapabilities.internetExplorer();
		capabilities.setCapability(InternetExplorerDriver.IGNORE_ZOOM_SETTING, true);
		System.setProperty("webdriver.ie.driver", CONFIG_IE_FILE.getAbsolutePath());
		this.DRIVER = new InternetExplorerDriver(capabilities);
		this.DRIVER.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS); //IMPLICIT WAIT 20
	}
	
	public InternetExplorerDriver getDriver()
	{
		return this.DRIVER;
	}
	
	public void setDriverExecutor()
	{
		this.EXECUTOR = (JavascriptExecutor)this.DRIVER;
	}
	
	public JavascriptExecutor getDriverExecutor()
	{
		return this.EXECUTOR;
	}

}
