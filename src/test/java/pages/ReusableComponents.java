package pages;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Random;
import java.util.Set;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import testBase.BaseClass;
import testBase.CommonActions;

public class ReusableComponents extends BaseClass {
	
	CommonActions common = new CommonActions();
	public By objLocator;
	Random random = new Random();
	String randomNum = String.valueOf(random.nextInt(100000));

	public void createList() throws Exception
	{		
		common.click("linkNew", "");
		common.click("linkNewList", "");	
		driver.switchTo().frame("createListFrame");
		Thread.sleep(3000);
		common.javascriptclick("linkBlankList", "");
		Thread.sleep(3000);
		common.javascriptclick("txtListName", "");		
		common.sendKeys("txtListName", "AutomationCheck", "");
		common.click("linkCheckbox", "");
		Thread.sleep(3000);
		common.click("btnCreate", "");
		Thread.sleep(3000);
		driver.switchTo().defaultContent();
		common.click("linkNSO", "");
		Thread.sleep(3000);
		common.click("linkSiteContents", "");
		Thread.sleep(3000);		
		common.click("txtSearchBox", "");
		common.sendKeys("txtSearchBox", "AutomationCheck", "");
		Robot robot=new Robot();
		robot.keyPress(KeyEvent.VK_ENTER);
		Thread.sleep(3000);
		common.click("linkSomeActions", "");
		Thread.sleep(3000);
		common.click("linkDelete", "");
		Thread.sleep(3000);		
		common.click("btnDelete", "");
		Thread.sleep(3000);
		common.click("linkNSO", "");
	
		}

	
	public void createDocumentLibrary() throws Exception
	{
		common.click("linkNew", "");
		common.click("linkDocumentLibrary", "");
		common.sendKeys("txtDocumentName", "AutomationCheck", "");
		common.click("linkCheckbox", "");
		Thread.sleep(3000);
		common.click("btnCreate", "");
		Thread.sleep(3000);
		common.click("btnNew", "");
		common.click("linkNewLink", "");
		common.sendKeys("txtLinkName", "www.google.com","");
		common.click("btnCreate", "");
		common.click("linkNSO", "");
		common.click("linkSiteContents", "");
		Thread.sleep(3000);		
		common.click("txtSearchBox", "");
		common.sendKeys("txtSearchBox", "AutomationCheck", "");
		Robot robot=new Robot();
		robot.keyPress(KeyEvent.VK_ENTER);
		Thread.sleep(3000);
		common.click("linkSomeActions", "");
		Thread.sleep(3000);
		common.click("linkDelete", "");
		Thread.sleep(3000);		
		common.click("btnDelete", "");
		Thread.sleep(3000);
		common.click("linkNSO", "");
	}
	
	
	public void createDocument() throws Exception
	{
		
		common.click("linkDocuments", "");
		Thread.sleep(5000);
		common.click("btnNew", "");
		Thread.sleep(5000);
		common.click("linkNewDocument", "");
		Thread.sleep(7000);
		writeDocument();
		driver.switchTo().defaultContent();
		common.click("linkNSO", "");
		Thread.sleep(3000);
		 
		
	   
	}
	public void writeDocument() throws AWTException, InterruptedException
	{
		Robot robot = new Robot();
		robot.keyPress(KeyEvent.VK_H);
		Thread.sleep(500);
		robot.keyPress(KeyEvent.VK_E);
		Thread.sleep(500);
		robot.keyPress(KeyEvent.VK_L);
		Thread.sleep(500);
		robot.keyPress(KeyEvent.VK_L);
		Thread.sleep(500);
		robot.keyPress(KeyEvent.VK_O);
		Thread.sleep(500);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_W);
	}
	
	
	public void uploadFunctionality() throws Exception
	{
		common.click("linkDocuments", "");
		Thread.sleep(3000);
		common.click("btnUpload", "");
		Thread.sleep(3000);	
		createNewFile(randomNum);
		Thread.sleep(5000);			
		common.click("linkFiles", "");
		Thread.sleep(3000);	
		setClipboardData(System.getProperty("user.dir")+"\\src\\test\\resources\\Assets\\Store\\NSO_Files\\New"+randomNum+".xlsx");
		uploadFileUsingRobotClass();
		Thread.sleep(7000);
		common.click("linkDocumentCheckbox", "");
		common.click("btnDownload", "");
		Thread.sleep(10000);
		Robot robot = new Robot();
		robot.keyPress(KeyEvent.VK_ENTER);
		common.click("linkDelete", "");
		common.click("btnDelete", "");
		Thread.sleep(3000);
		common.click("linkNSO", "");
		Thread.sleep(3000);
	}
	
	
	public void createNewFile(String newFile) throws Exception{
		CopyArchiveFile("New.xlsx","New"+newFile+".xlsx");
		
	}
	 public void CopyArchiveFile(String sourceFileName, String destFileName)
	    {
	    	File source = new File(System.getProperty("user.dir")+"\\src\\test\\resources\\Assets\\Store\\"+sourceFileName);
	    	File dest = new File(System.getProperty("user.dir")+"\\src\\test\\resources\\Assets\\Store\\NSO_Files\\"+destFileName);
	    	try {
	    	    FileUtils.copyFile(source, dest);
	    	    
	    	} catch (IOException e) {
	    	    e.printStackTrace();
	    	}
	    }
	 
	 public static void setClipboardData(String path) {
			StringSelection stringSelection = new StringSelection(path);
			Toolkit.getDefaultToolkit().getSystemClipboard().setContents(stringSelection, null);
		}
		
	 public void uploadFileUsingRobotClass() throws AWTException, InterruptedException
	    {
	    	//native key strokes for CTRL, V and ENTER keys
			Robot robot = new Robot();
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_V);
			robot.keyRelease(KeyEvent.VK_V);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_ENTER);
			robot.keyRelease(KeyEvent.VK_ENTER);
			Thread.sleep(1000);
			
	    }
	 
	 
	public boolean elementExists(String objectName, String RuntimeLocator) throws Exception {
		try {
			objLocator = getElementLocator(objectName, RuntimeLocator);
			WebDriverWait wait = new WebDriverWait(driver, 30);
			wait.until(ExpectedConditions.presenceOfElementLocated(objLocator));
			if (driver.findElement(objLocator).isDisplayed()) {
				// WebElement objclick = driver.findElement(objLocator);
				// objclick.click();
				logged.info("Element is visible " + objectName);
				logResult("Visible " + objectName, "Element should be visible " + objectName,
						objectName + " is visible", "Pass", "");
				return true;
			} else {
				String screenshotPath = captureScreenshot("Visible " + objectName);
				logged.error("Element is not visible" + objectName);
				logResult("Visible " + objectName, "Element should be visible " + objectName,
						"Element is not visible " + objectName, "Fail", screenshotPath);
				return false;
			}
		} catch (Exception clickException) {
			String screenshotPath = captureScreenshot("Visible " + objectName);
			logged.error("Element is not visible " + objectName);
			logResult("Visible " + objectName, "Element should be visible" + objectName,
					"Element is not visible " + objectName, "Fail", screenshotPath);
			return false;
			
		}
	}
}
