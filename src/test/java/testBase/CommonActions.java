package testBase;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class CommonActions extends BaseClass{
	
	public By objLocator;
	
	public CommonActions() {
		PageFactory.initElements(driver, this);
	}
	
	/**
	This method is used to invoke the browser provided in Configuration properties file
	*/	
	public void invokeBrowser() throws InterruptedException {
		
		String browsername = prop.getProperty("Browser");
		if(browsername.equalsIgnoreCase("edge")) {
			System.setProperty("webdriver.edge.driver","./src/test/resources/Assets/BrowserDriver/msedgedriver.exe");			
			driver  = new EdgeDriver();
		}else if(browsername.equalsIgnoreCase("chrome")) {
			System.setProperty("webdriver.chrome.driver","./src/test/resources/Assets/BrowserDriver/chromedriver.exe");
			driver  = new ChromeDriver();
		}
		else if(browsername.equalsIgnoreCase("ie")) {
			System.setProperty("webdriver.ie.driver","./src/test/resources/Assets/BrowserDriver/IEDriverServer.exe");
			driver  = new InternetExplorerDriver();
		}
		driver.manage().window().maximize();
		driver.manage().deleteAllCookies();
		driver.manage().timeouts().pageLoadTimeout(30, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);				
		
	}

	/**
	This method is used to set the URL according to the environment provided in excel
	*/								
    public void setEnvtURL() throws InterruptedException, IOException {
    	
    	String envtName = (String)lstEnvtsheet.get(0);
		if(envtName.equalsIgnoreCase("qa")) {
			driver.get(prop.getProperty("qaurl"));			
		}else if(envtName.equalsIgnoreCase("test")) {
			driver.get(prop.getProperty("testurl"));
		}
		else if(envtName.equalsIgnoreCase("dev")) {
			driver.get(prop.getProperty("devurl"));
		}
		else if(envtName.equalsIgnoreCase("dev2")) {
			driver.get(prop.getProperty("dev2temp"));
		}
		driver.manage().timeouts().pageLoadTimeout(30, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);						
	}
   
    /**
	This method is used to close the browser session created
	*/
	public void closeBrowser() {
		
		driver.quit();
		
	}
	
	/**
	This method is used to click the object
	@param objectName - object name given in excel object repository
	@param RuntimeLocator - value to be fetched at runtime else give as null("") 
	@return boolean - true(if click is done) or else false
	*/
	public boolean click(String objectName,String RuntimeLocator) throws Exception		    
	  {
	    try
	    {	      
	      objLocator = getElementLocator(objectName,RuntimeLocator);
	      WebDriverWait wait = new WebDriverWait(driver,30);
	         wait.until(ExpectedConditions.presenceOfElementLocated(objLocator));
	      if(driver.findElement(objLocator).isEnabled()) {
	    	  WebElement objclick = driver.findElement(objLocator);
	    	  objclick.click();
	    	  logged.info("User is able to click "+objectName);
	    	  logResult("click "+objectName, "User should be able to click "+objectName, objectName+" is clicked successfully", "Pass", "");			  	        
			  return true;
	      } else {
	    	  String screenshotPath = captureScreenshot("click "+objectName);
	    	  logged.error("User is not able to click "+objectName);
	    	  logResult("click "+objectName,"User should be able to click "+objectName, "Not able to click "+objectName, "Fail", screenshotPath);	    	  
	    	  return false;
	      }
	    } catch (Exception clickException) {
	    	String screenshotPath = captureScreenshot("click "+objectName);
	    	logged.error("User is not able to click "+objectName);
	    	logResult("click "+objectName,"User should be able to click "+objectName, "Not able to click "+objectName, "Fail", screenshotPath);	    	
	    	return false;
	    }
	 }
	
	/**
	This method is used to enter the text for the object
	@param objectName - object name given in excel object repository
	@param strText - text which needs to be entered 
	@param RuntimeLocator - value to be fetched at runtime else give as null("") 
	@return boolean - true(if text is entered) or else false
	*/
	public boolean sendKeys(String objectName,String strText,String RuntimeLocator) throws Exception		    
	  {
	    try
	    {	      
	      objLocator = getElementLocator(objectName,RuntimeLocator);
	      WebDriverWait wait = new WebDriverWait(driver,30);
	         wait.until(ExpectedConditions.presenceOfElementLocated(objLocator));
	      if(driver.findElement(objLocator).isEnabled()) {
	    	  WebElement objsendKeys = driver.findElement(objLocator);
	    	  objsendKeys.sendKeys(strText);
	    	  logged.info("User is able to enter text"+strText+" for "+objectName);
	    	  logResult("Enter text for "+objectName, "User should be able to enter text "+strText+" for "+objectName, "Text "+strText+" is entered successfully for "+objectName, "Pass", "");			  	        
			  return true;
	      } else {
	    	  String screenshotPath = captureScreenshot("Enter text for "+objectName);
	    	  logged.error("User is not able to enter text"+strText+" for "+objectName);
	    	  logResult("Enter text for "+objectName, "User should be able to enter text "+strText+" for "+objectName, "Unable to enter text "+strText+" for "+objectName, "Fail", screenshotPath);	    	  
	    	  return false;
	      }
	    } catch (Exception sendKeysException) {
	    	String screenshotPath = captureScreenshot("Enter text for "+objectName);
	    	logged.error("User is not able to enter text"+strText+" for "+objectName);	    	
	    	logResult("Enter text for "+objectName, "User should be able to enter text "+strText+" for "+objectName, "Unable to enter text "+strText+" for "+objectName, "Fail", screenshotPath);
	    	return false;
	    }
	 }
	
	/**
	This method is used to select the text for the object
	@param ObjectName - object name given in excel object repository
	@param strText - text which needs to be selected 
	@param RuntimeLocator - value to be fetched at runtime else give as null("") 
	@return boolean - true(if text is selected) or else false
	*/
	public boolean selectText(String objectName,String strText,String RuntimeLocator) throws Exception		    
	  {
	    try
	    {	      
	      objLocator = getElementLocator(objectName,RuntimeLocator);
	      WebDriverWait wait = new WebDriverWait(driver,30);
	         wait.until(ExpectedConditions.presenceOfElementLocated(objLocator));
	      if(driver.findElement(objLocator).isEnabled()) {
	    	  WebElement objSelect = driver.findElement(objLocator);
	    	  Select select = new Select(objSelect);
	    	  select.selectByVisibleText(strText);
	    	  logged.info("User is able to select text"+strText+" for "+objectName);
	    	  logResult("Select text for "+objectName, "User should be able to select text "+strText+" for "+objectName, "Text "+strText+" is selected successfully for "+objectName, "Pass", "");			          
			  return true;
	      } else {
	    	  String screenshotPath = captureScreenshot("Select text for "+objectName);
	    	  logged.error("User is not able to select text"+strText+" for "+objectName);
	    	  logResult("Select text for "+objectName, "User should be able to select text "+strText+" for "+objectName, "Unable to select text "+strText+" for "+objectName, "Fail",screenshotPath);	    	  
	    	  return false;
	      }
	    } catch (Exception selectTextException) {
	    	String screenshotPath = captureScreenshot("Select text for "+objectName);
	    	logged.error("User is not able to select text"+strText+" for "+objectName);
	    	logResult("Select text for "+objectName, "User should be able to select text "+strText+" for "+objectName, "Unable to select text "+strText+" for "+objectName, "Fail", screenshotPath);	    	
	    	return false;
	    }	        
	 }
	
	
	/**
	This method is used to click the object by javascript executor
	@param objectName - object name given in excel object repository
	@param RuntimeLocator - value to be fetched at runtime else give as null("") 
	@return null
	*/
	public void javascriptclick(String objectName,String RuntimeLocator) throws Exception		    
	  {
	    try
	    {	      
	      objLocator = getElementLocator(objectName,RuntimeLocator);
	      WebDriverWait wait = new WebDriverWait(driver,30);
	         wait.until(ExpectedConditions.presenceOfElementLocated(objLocator));
	      if(driver.findElement(objLocator).isEnabled()) {
	    	  WebElement objclick = driver.findElement(objLocator);
	    	  JavascriptExecutor executor = (JavascriptExecutor)driver;
	          executor.executeScript("arguments[0].click();", objclick);
	    	  logged.info("User is able to click "+objectName);
	    	  logResult("Click "+objectName, "User should be able to click "+objectName,  objectName+" is clicked successfully", "Pass", "");			          
	      } else {
	    	  String screenshotPath = captureScreenshot("Click "+objectName);
	    	  logged.error("User is not able to click "+objectName);
	    	  logResult("Click "+objectName, "User should be able to click "+objectName, "Not able to click "+objectName, "Fail", screenshotPath);	    	  
	      }
	    } catch (Exception clickException) {
	    	String screenshotPath = captureScreenshot("Click "+objectName);
	    	logged.error("User is not able to click "+objectName);
	    	logResult("Click "+objectName, "User should be able to click "+objectName, "Not able to click "+objectName, "Fail", screenshotPath);	    	
	    }
	 }
	
	/**
	This method is used to hover on the object
	@param objectName - object name given in excel object repository
	@param RuntimeLocator - value to be fetched at runtime else give as null("") 
	@return null
	*/
	public void hover(String objectName,String RuntimeLocator) throws Exception		    
	  {
	    try
	    {	      
	      objLocator = getElementLocator(objectName,RuntimeLocator);
	      Actions hover = new Actions(driver);
	      WebDriverWait wait = new WebDriverWait(driver,30);
	         wait.until(ExpectedConditions.presenceOfElementLocated(objLocator));
	      if(driver.findElement(objLocator).isEnabled()) {
	    	  WebElement objhover = driver.findElement(objLocator);
	    	  hover.moveToElement((WebElement)objhover).build().perform();
	    	  logged.info("User is able to hover on "+objectName);
	    	  logResult("Hover on "+objectName, "User should be able to hover on "+objectName, objectName+" is hovered successfully", "Pass", "");	        
	      } else {
	    	  String screenshotPath = captureScreenshot("Hover on "+objectName);
	    	  logged.error("User is not able to hover on "+objectName);
	    	  logResult("Hover on "+objectName, "User should be able to hover on "+objectName,"Not able to hover on "+objectName, "Fail", screenshotPath);	    	  
	      }
	    } catch (Exception hoverException) {
	    	String screenshotPath = captureScreenshot("Hover on "+objectName);
	    	logged.error("User is not able to hover on "+objectName);
	    	logResult("Hover on "+objectName, "User should be able to hover on "+objectName, "Not able to hover on "+objectName, "Fail", screenshotPath);	    	
	    }
	 }
	
	/**
	This method is used to hover and click on the object
	@param objectName - object name given in excel object repository
	@param RuntimeLocator - value to be fetched at runtime else give as null("") 
	@return null
	*/
	public void hoverClick(String objectName,String RuntimeLocator) throws Exception		    
	  {
	    try
	    {	      
	      objLocator = getElementLocator(objectName,RuntimeLocator);
	      Actions hover = new Actions(driver);
	      WebDriverWait wait = new WebDriverWait(driver,30);
	         wait.until(ExpectedConditions.presenceOfElementLocated(objLocator));
	      if(driver.findElement(objLocator).isEnabled()) {
	    	  WebElement objhover = driver.findElement(objLocator);
	    	  hover.moveToElement((WebElement)objhover).click((WebElement)objhover).build().perform();
	    	  logged.info("User is able to hover and click "+objectName);
	    	  logResult("Hover and click "+objectName, "User should be able to hover and click "+objectName, objectName+" is hovered and clicked successfully", "Pass", "");			  	        
	      } else {
	    	  String screenshotPath = captureScreenshot("Hover and click "+objectName);
	    	  logged.error("User is not able to hover and click "+objectName);
	    	  logResult("Hover and click "+objectName, "User should be able to hover and click "+objectName, "Not able to hover  and click "+objectName, "Fail", screenshotPath);	    	  
	      }
	    } catch (Exception hoverException) {
	    	String screenshotPath = captureScreenshot("Hover and click "+objectName);
	    	logged.error("User is not able to hover and click "+objectName);
	    	logResult("Hover and click "+objectName, "User should be able to hover and click "+objectName, "Not able to hover  and click "+objectName, "Fail", screenshotPath);	    	
	    }
	 }
	
	/**
	This method is used to verify object text and given text
	@param objectName - object name given in excel object repository
	@param strExpectedText - name given in test data which contains expected text
	@param RuntimeLocator - value to be fetched at runtime else give as null("") 
	@return true(if text matches) or else false
	@modifiedby maballa on 21 Nov 2019
	*/
	public boolean verifyText(String objectName, String strExpectedText, boolean blnRemoveSapces, boolean blnIgnoreCase, String RuntimeLocator) throws Exception		    
	  {
		String strActualText = "";
	    try
	    {
	    	objLocator = getElementLocator(objectName,RuntimeLocator);
	    	WebDriverWait wait = new WebDriverWait(driver,30);
	         wait.until(ExpectedConditions.presenceOfElementLocated(objLocator));
	        if(driver.findElement(objLocator).isEnabled()) {
		    	  WebElement objtext = driver.findElement(objLocator);
		    	  strActualText = objtext.getText();
		    			  
		    	  if(blnRemoveSapces==true) {
			    	  strActualText = strActualText.replace(" ","");
  			  		  strExpectedText = strExpectedText.replace(" ","");
		    	  }
		    	  if(blnIgnoreCase ==true) {
			    	  strActualText = strActualText.toUpperCase();
  			  		  strExpectedText = strExpectedText.toUpperCase();
		    	  }
		    	   
		    	  if (strExpectedText.equals(strActualText)) {
			    	  logged.info(objectName+" Expected text value "+strExpectedText+" is same as Actual text value "+strActualText);
			    	  logResult("verify text displayed for "+objectName, "Text should be displayed as "+strExpectedText+" for "+objectName,objectName+" Expected text value "+strExpectedText+" is same as Actual text value "+strActualText, "Pass", "");	 			  	        
		 			  return true;
		    	  } 
		    	  else {
		    		  String screenshotPath = captureScreenshot("verify text displayed for "+objectName);
			    	  logged.error(objectName+" Expected text value "+strExpectedText+" is different from Actual text value "+strActualText);
			    	  logResult("verify text displayed for "+objectName, "Text should be displayed as "+strExpectedText+" for "+objectName,objectName+" Expected text value "+strExpectedText+" is different from Actual text value "+strActualText, "Fail", screenshotPath);			    	  
			    	  return false;
		    	  }
	        }
		    else {		
		    	  String screenshotPath = captureScreenshot("verify text displayed for "+objectName);
		    	  logged.error("Not able to find "+objectName);
		    	  logResult("verify text displayed for "+objectName, "Text should be displayed as "+strExpectedText+" for "+objectName,"Not able to find "+objectName, "Fail", screenshotPath);		    	  
		    	  return false;			    	 
		    	  }	    	 
	    } catch (Exception verifyTextException) {
	    	String screenshotPath = captureScreenshot("verify text displayed for "+objectName);
	    	logged.error("Not able to find "+objectName);
	    	logResult("verify text displayed for "+objectName, "Text should be displayed as "+strExpectedText+" for "+objectName,"Not able to find "+objectName, "Fail", screenshotPath);	    	
	    	return false;
	    }
	 }
	
	/**
	This method is used to wait until frame gets available and switch to it
	@param framelocator - frame name given in excel object repository
	@param RuntimeLocator - value to be fetched at runtime else give as null("") 	
	@return null
	*/
	 public void waitAndSwitchToFrame(String frameName,String RuntimeLocator) throws Exception			    
	   {
		 try {
			 By frameLocator = getElementLocator(frameName,RuntimeLocator);
			 WebDriverWait switchFrame = new WebDriverWait(driver, 30);
			 	switchFrame.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt(frameLocator)); 		
		 	 logged.info("User is able to switch to frame "+frameName);
		 	logResult("Switch to frame"+frameName,  "User should be able to switch to frame "+frameName,"User is able to switch to frame "+frameName+" successfully", "Pass", "");			 		 
		 }
		 catch (Exception switchFrameException) {
			 String screenshotPath = captureScreenshot("Switch to frame"+frameName);
	    	 logged.error("User is not able to switch to frame"+frameName);
	    	 logResult("Switch to frame"+frameName, "User should be able to switch to frame "+frameName,"User is not able to switch to frame"+frameName, "Fail", screenshotPath);		    	
		 }
	   }
	 
	 /**
		This method is used to clear the object text
		@param objectName - object name given in excel object repository
		@param RuntimeLocator - value to be fetched at runtime else give as null("") 
		@return null
		*/
	public void clear(String objectName,String RuntimeLocator) throws Exception		    
	  {
	    try
	    {	      
	      objLocator = getElementLocator(objectName,RuntimeLocator);
	      WebDriverWait wait = new WebDriverWait(driver,30);
	         wait.until(ExpectedConditions.presenceOfElementLocated(objLocator));
	      if(driver.findElement(objLocator).isEnabled()) {
	    	  WebElement objclear = driver.findElement(objLocator);
	    	  objclear.clear();
	    	  logged.info("User is able to clear text from "+objectName);
	    	  logResult("Clear text from "+objectName, "User should be able to clear text from "+objectName, objectName+" text is cleared successfully", "Pass", "");  			  	        
	      } else {
	    	  String screenshotPath = captureScreenshot("Clear text from "+objectName);
	    	  logged.error("User is not able to text from "+objectName);
	    	  logResult("Clear text from "+objectName, "User should be able to clear text from "+objectName,"Not able to clear text from "+objectName, "Fail", screenshotPath);	    	  
	      }
	    } catch (Exception clearException) {
	    	String screenshotPath = captureScreenshot("Clear text from "+objectName);
	    	logged.error("User is not able to text from "+objectName);
	    	logResult("Clear text from "+objectName, "User should be able to clear text from "+objectName,"Not able to clear text from "+objectName, "Fail", screenshotPath);	    	
	    }
	 }
		
	/**
	This method is used to verify alert exists and accept the alert
	*/
	public void acceptAlert() throws Exception		    
	  {
	    try
	    {	      		      
	      WebDriverWait wait = new WebDriverWait(driver,30);
	         wait.until(ExpectedConditions.alertIsPresent());
	      Alert alert = driver.switchTo().alert();
	      alert.accept();
    	  logged.info("User is able to accept alert");
    	  logResult("Accept alert", "User should be able to accept alert","User is able to accept alert successfully", "Pass", "");		  	        		      
	    } 
	    catch (Exception acceptAlertException) {
		   String screenshotPath = captureScreenshot("Accept alert");
		   logged.error("User is not able to accept alert");
		   logResult("Accept alert", "User should be able to accept alert","User is not able to accept alert", "Fail", screenshotPath);	      
	    }
	 }
	
	/**
	This method is used to verify alert exists and dismiss the alert
	*/
	public void dismissAlert() throws Exception		    
	  {
	    try
	    {	      		      
	      WebDriverWait wait = new WebDriverWait(driver,30);
	         wait.until(ExpectedConditions.alertIsPresent());
	      Alert alert = driver.switchTo().alert();
	      alert.dismiss();
    	  logged.info("User is able to dismiss alert");
    	  logResult("Dismiss alert", "User should be able to dismiss alert","User is able to dismiss alert successfully", "Pass", "");		          		      
	    } 
	    catch (Exception dismissAlertException) {
	       String screenshotPath = captureScreenshot("Dismiss alert");
	       logged.error("User is not able to dismiss alert");
	       logResult("Dismiss alert", "User should be able to dismiss alert","User is not able to dismiss alert", "Fail", screenshotPath);
	      
	    }
	 }
	
	/**
	This method is used to verify alert exists and get the text displayed in alert
	@return text message displayed in the alert
	*/
	public String getAlertText() throws Exception		    
	  {
		String alertTextMsg = "";
	    try
	    {	      		      
	      WebDriverWait wait = new WebDriverWait(driver,30);
	         wait.until(ExpectedConditions.alertIsPresent());
	      Alert alert = driver.switchTo().alert();
	      alertTextMsg = alert.getText();
    	  logged.info(alertTextMsg+" message is displayed in alert");
    	  logResult("Get Alert Text Message", "User should be able to get the text displayed in alert",alertTextMsg+" message is displayed in alert", "Pass", "");		  	        		      
		  return alertTextMsg; 
	    } 
	    catch (Exception getAlertTextException) {
	       String screenshotPath = captureScreenshot("Get Alert Text Message");
	       logged.error("User is not able to get the text displayed in alert");
	       logResult("Get Alert Text Message", "User should be able to get the text displayed in alert","User is not able to get the text displayed in alert", "Fail", screenshotPath);	      
	    }
	    return alertTextMsg;
	 }
	
	/**
	This method is used to verify alert exists and set the text to be displayed in alert
	@param text message to be entered on the alert
	*/
	public void setAlertText(String alertTextMsg) throws Exception		    
	  {
	    try
	    {	      		      
	      WebDriverWait wait = new WebDriverWait(driver,30);
	         wait.until(ExpectedConditions.alertIsPresent());
	      Alert alert = driver.switchTo().alert();
	      alert.sendKeys(alertTextMsg);
    	  logged.info("User is able to enter "+alertTextMsg+" message on alert");
    	  logResult("Set Alert Text Message", "User should be able to enter "+alertTextMsg+" message on alert","User is able to enter "+alertTextMsg+" message on alert successfully", "Pass", "");		  	        		       
	    } 
	    catch (Exception setAlertTextException) {
	    	String screenshotPath = captureScreenshot("Set Alert Text Message");
	        logged.error("User is not able to enter "+alertTextMsg+" message on alert");
	        logResult("Set Alert Text Message", "User should be able to enter "+alertTextMsg+" message on alert","User is not able to enter "+alertTextMsg+" message on alert", "Fail", screenshotPath);	       
	    }
	 }
	
	
	/**
	This method is used to write the data to excel
	@param excel filepath 
	@param excelname 
	@param sheetname where we want data to be written in excel
	@param strValue is the value which we want to be written in given rowno and colno
	*/
	public void writeExcel(String filePath, String excelName, String sheetName, int rowno, int colno, String strValue)
		      throws IOException
	  {
	     try
	     {	    	
	        String excelfilepath = filePath + "\\" + excelName;	        
	        File file = new File(excelfilepath);       
	        if (file.exists()) {
	          FileInputStream inputStream = new FileInputStream(file);
	          XSSFWorkbook Workbook1 = new XSSFWorkbook(inputStream);         
	          Sheet sheet = Workbook1.getSheet(sheetName);          
	          Row row = sheet.createRow(rowno);
	          Cell cell = row.createCell(colno);
	          cell.setCellType(1);
	          cell.setCellValue(strValue);          
	          inputStream.close();          
	          FileOutputStream outputStream = new FileOutputStream(file);         
	          Workbook1.write(outputStream);          
	          outputStream.close();
	          Workbook1.close();
	          logged.info("User is able to write data "+strValue+" in excel");
	          logResult("Write Data in excel", "User should be able to write data "+strValue+" in excel","User is able to write data "+strValue+" in excel successfully", "Pass", "");			  	        		       
		    } 
	        else {
	          String screenshotPath = captureScreenshot("Write Data in excel");
	  	      logged.error("User is not able to write data "+strValue+" in excel");
	  	      logResult("Write Data in excel", "User should be able to write data "+strValue+" in excel","User is not able to write data "+strValue+" in excel", "Fail", screenshotPath);	  	      
	  	    }
	     }catch (Exception writeExcelException) {
	      String screenshotPath = captureScreenshot("Write Data in excel");
	      logged.error("User is not able to write data "+strValue+" in excel");
	      logResult("Write Data in excel", "User should be able to write data "+strValue+" in excel","User is not able to write data "+strValue+" in excel", "Fail", screenshotPath);	      
	    }
          
	 }
	    
	/**
	This method is used to read the data from excel
	@param excel filepath 
	@param excelname 
	@param sheetname where we want data to be written in excel	
	@return data read from given rowno and colno in excel
	*/	
    public String readExcel(String excelFilePath, String sheetName, int rowno, int colno)
      throws IOException
    {
      String strExcelValue = "";
      try {   	        	         
          FileInputStream inputStream = new FileInputStream(excelFilePath);
          Workbook workbook = new XSSFWorkbook(inputStream);
          Sheet sheet = workbook.getSheet(sheetName);  
          Row row = sheet.getRow(rowno);
          Cell cell = row.getCell(colno);
          strExcelValue = cell.getStringCellValue();
          workbook.close();
          inputStream.close();         
          logged.info("User is able to read data "+strExcelValue+" from excel");
          logResult("Read Data from excel", "User should be able to read data "+strExcelValue+" from excel","User is able to read data "+strExcelValue+" from excel successfully", "Pass","");		  		       		   
      }
      catch (Exception readExcelException) {
	      String screenshotPath = captureScreenshot("Read Data from excel");
	      logged.error("User is not able to read data "+strExcelValue+" from excel");
	      logResult("Read Data from excel", "User should be able to read data "+strExcelValue+" from excel","User is not able to read data "+strExcelValue+" from excel", "Fail", screenshotPath);
     }         
      return strExcelValue;
    }
	
    /**
	This method is used to scroll horizontally
	@param objectName - object name given in excel object repository 
	@param RuntimeLocator - value to be fetched at runtime else give as null("") 	
	*/
    public void scrollHorizontal(String objectName,String RuntimeLocator) throws Exception   	      
     {
	  try {      
	      objLocator = getElementLocator(objectName,RuntimeLocator);
	      WebDriverWait wait = new WebDriverWait(driver,30);
	         wait.until(ExpectedConditions.elementToBeClickable(objLocator));       
	      JavascriptExecutor executor = (JavascriptExecutor)driver;
	      executor.executeScript("arguments[0].scrollIntoView();", new Object[] { objLocator });
	      logged.info("User is able to scroll horizontally");
	      logResult("Scroll horizontally", "User should be able to scroll horizontally","User is able to scroll horizontally", "Pass", "");		  	                 	  
	    } catch (Exception scrollHorizontalException) {
	    	String screenshotPath = captureScreenshot("Scroll horizontally");
	    	logged.error("User is not able to scroll horizontally");
	    	logResult("Scroll horizontally", "User should be able to scroll horizontally","User is not able to scroll horizontally", "Fail", screenshotPath);
	    	
	      } 
     }
    
	/**
	 * This method generates a random number
	 * 
	 * @return double
	 * @author subhamohapatra
	 * @throws Exception
	 * @modifiedby maballa on 21 Nov 2019
	 */
	public static int getRandomIntegerBetweenRange(double min, double max) {
		int x = (int) ( Math.random() * ((max - min) + 1) + min);
		return x;
	}


  /**
   	This method is used to send keys of users choice to a field
   	@param objectName - object name given in excel object repository 
   	@param strKey - Key which needs to be entered 
	@param RuntimeLocator - value to be fetched at runtime else give as null("") 
	@return boolean - true(if text is entered) or else false
	@author shanandakumar
	*/
    public boolean sendkeysCustom(String objectName,String strKey,String RuntimeLocator) throws IOException
	{
		try
	    {	      
	      objLocator = getElementLocator(objectName,RuntimeLocator);
	      WebDriverWait wait = new WebDriverWait(driver,30);
	         wait.until(ExpectedConditions.presenceOfElementLocated(objLocator));
	      if(driver.findElement(objLocator).isEnabled()) {
	    	  if(strKey.equalsIgnoreCase("tab"))
	    		{
	    		  WebElement objsendKeys = driver.findElement(objLocator);
	    		  objsendKeys.sendKeys(Keys.TAB);
	    		}
	    	  else if(strKey.equalsIgnoreCase("enter"))
	    	  	{
	    		  WebElement objsendKeys = driver.findElement(objLocator);
	    		  objsendKeys.sendKeys(Keys.ENTER);
	    		}
	    	  logged.info("User is able to select text"+strKey+" for "+objectName);
	    	  logResult("Select text for "+objectName, "User should be able to select "+strKey+" for "+objectName, "Key "+strKey+" is selected successfully for "+objectName, "Pass", "");			          
			  return true;
	      } else {
	    	  String screenshotPath = captureScreenshot("Select text for "+objectName);
	    	  logged.error("User is not able to select "+strKey+" for "+objectName);
	    	  logResult("Select text for "+objectName, "User should be able to select "+strKey+" for "+objectName, "Unable to select "+strKey+" for "+objectName, "Fail",screenshotPath);	    	  
	    	  return false;
	      }
	    } catch (Exception selectTextException) {
	    	String screenshotPath = captureScreenshot("Enter text for "+objectName);
	    	logged.error("User is not able to enter "+strKey+" for "+objectName);	    	
	    	logResult("Enter text for "+objectName, "User should be able to enter "+strKey+" for "+objectName, "Unable to enter "+strKey+" for "+objectName, "Fail", screenshotPath);
	    	return false;
	    }

	}

    /**
   	This method is used to enter HTML report step based on pass/fail 
   	@param strStepStatus: pass/fail
   	@param strAction: Validation action step 
	@param strExpected: Expected result  
	@param strActual: Actual result
	@param ScreenShotNamescreen boolean - true(if text is entered) or else false
	@author maballa
	*/
	public void HTMLReportStmt(String strStepStatus, String strAction, String strExpected, String strActual, String ScreenShotName) throws Exception{
		if(strStepStatus.toUpperCase().contentEquals("PASS")) {
			logged.info(strActual);
			logResult(strAction, strExpected, strActual, "pass", "");
		}else {
			logged.info(strActual);
			String strScreenshotPath = captureScreenshot(ScreenShotName);
			logResult(strAction, strExpected, strActual, "fail", strScreenshotPath);
		}
	}
	}
