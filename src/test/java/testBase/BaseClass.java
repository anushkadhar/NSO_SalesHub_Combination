package testBase;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import org.apache.commons.io.FileUtils;
import org.apache.log4j.FileAppender;
import org.apache.log4j.Layout;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.log4j.spi.ErrorCode;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.testng.SkipException;
import org.testng.TestNG;
import org.testng.xml.XmlClass;
import org.testng.xml.XmlInclude;
import org.testng.xml.XmlSuite;
import org.testng.xml.XmlTest;

public class BaseClass extends FileAppender{
	
	  public static List<String> lstTestExecutionsheet;
	  public static List<String> lstEnvtsheet;
	  public static List<String> lstLocators;
	  public static Map<String, String> mapTestData = new HashMap();
	  public static List listkeyValue = new ArrayList();
	  public static Map<Integer, List<String>> mapTestExecution = new HashMap();
	  public static Map<String, List<String>> mapObjectRepository = new HashMap();
	  public static By ElementLocator;
	  public static By ElementLocate;
	  public static Sheet EnvtDetails = null;
	  public static Sheet TestExecutionDetails = null;	  
	  public static StringBuffer sbDetailedReport = new StringBuffer();
	  public static int intSNo = 1;
	  public static int Sno = 1;
	  public static int passStepsCount = 0;
	  public static int failStepsCount = 0;
	  public static String strResultStatus = "";
	  public static String strReportsFolderPath = null;
	  public static String strExecutionDate = null;
	  public static String IEPath;
	  public static WebDriver driver;
	  public static Properties prop;	
	  public Logger logged = Logger.getLogger("log statements");
	  public String logPropertyPath= "./src/test/resources/Assets/Configuration/log4j.properties";
	  public String propertyFilePath= "./src/test/resources/Assets/Configuration/Config.properties";	  
	  public String excelFilePath = "./src/test/resources/Assets/testData/AutomationTestSuite.xlsx";  
	  public String objectRepositoryFilePath = "./src/test/resources/Assets/objectRepository/ObjectRepository.xlsx";
	  public static int iHours = 0;
	  public static int iSeconds = 0;
	  public static int iMinutes = 0;
	  public static int intTcsExecutedCount = 0;
	  public static String strExecutionStatus = "";
	  public static String strSummaryReportName = "";
	
	public BaseClass() {
    }
    
    public BaseClass(Layout layout, String filename,
    		boolean append, boolean bufferedIO, int bufferSize)
    		throws IOException {
    	super(layout, filename, append, bufferedIO, bufferSize);
    }
    
    public BaseClass(Layout layout, String filename,
    		boolean append) throws IOException {
    	super(layout, filename, append);
    }
    
    public BaseClass(Layout layout, String filename)
    		throws IOException {
    	super(layout, filename);
    }

	  
  //******************** Log4j methods ************
    
    /**
	This method is used to set the file options for log4j properties file
	*/	
    
	public void activateOptions() {
	    if (fileName != null) {
	    	try {
	    		fileName = getNewLogFileName();
	    		setFile(fileName, fileAppend, bufferedIO, bufferSize);
	    	} catch (Exception e) {
	    		errorHandler.error("Error while activating log options", e,
	    				ErrorCode.FILE_OPEN_FAILURE);
	    	}
	    }
	}
	 
	/**
	This method is used to get the log4j file name with current timestamp
	@return log4jfilename
	*/
	 private String getNewLogFileName() {
		 
	    if (fileName != null) {
	    	final File logFile = new File(fileName);
    		final String fileName[] = (logFile.getName()).split("\\.");
    		String newFileName = "";			    
    		String strLogDate = getCurrentDateTime("ddMMMyyyy")+"_"+getCurrentDateTime("HHmmss");
    		newFileName = fileName[0]+"_"+strLogDate+"."+fileName[1];	
	    	return logFile.getParent()+File.separator+newFileName;
	    }
	    return null;
	  }	
	
	  
	//******************** Test Execution Flow methods ************
	  
	 /**
		This method is used to read the configuration properties file and store values
		@return configParameters - Properties defined in configuration properties file
		*/
	  public String[] readConfigurationProperties() throws Exception {
			    
	    try
	    {
	      Properties prop = new Properties();
	      String[] configParameters = new String[14];	      
	      InputStream fileip = new FileInputStream(propertyFilePath);	      
	      if (fileip != null) {
	        prop.load(fileip);
	        String strBrowser = prop.getProperty("Browser");	       	        	         
	        if (strBrowser == null) {
	            throw new SkipException("Skipping execution - Issue with Configuration Property File - Browser parameter is missing");
	          }	          
	        else {
	          if ((strBrowser.contentEquals("")) )
	              throw new SkipException("Skipping execution - Issue with Configuration Property File - Browser parameter does not contain any value");	       
	          }	          
	          configParameters[0] = strBrowser;	
	        }	      
	      return configParameters;
	    }
	    catch (Exception readConfigurationPropertiesException) {
	      System.out.println("readCoonfigurationProperties Method Exception Message: " + readConfigurationPropertiesException);
	      throw new SkipException("Skipping execution - Issue with Configuration Property File");
		    }
		  }
		
	  /**
		This method is used to read the Environment Details from the testdata excel sheet
		*/
	  public void readEnvtDetails()  throws IOException {
	   		
		    try
		    {
		      FileInputStream inputStream = new FileInputStream(excelFilePath);	      
		      Workbook workbook = new XSSFWorkbook(inputStream);
		      EnvtDetails = workbook.getSheetAt(0); 	      
		      for (int rowno = 0; rowno < EnvtDetails.getLastRowNum(); rowno++) {
		          Row nextRow = EnvtDetails.getRow(rowno + 1);
		          Cell envtName = nextRow.getCell(0);	         
		          String strEnvtName = envtName.getStringCellValue();
		          Cell suiteName = nextRow.getCell(1);
		          String strSuiteName = suiteName.getStringCellValue(); 
		          Cell projectName = nextRow.getCell(2);
		          String strProjectName = projectName.getStringCellValue(); 
		          lstEnvtsheet = new ArrayList();
		          lstEnvtsheet.add(strEnvtName);
		          lstEnvtsheet.add(strSuiteName);
		          lstEnvtsheet.add(strProjectName);
		      }
	        workbook.close();
	        inputStream.close();
	    } catch (Exception readEnvtDetailsException) {
	      System.out.println("readEnvtDetails Method Exception Message: " + readEnvtDetailsException);
	      throw new SkipException("Skipping execution - Issue with Environment Details Sheet");
	    }
	  }
	  
		
	  /**
		This method is used to read the Test Execution Details from the testdata excel sheet
		*/
	  public void readTestExecution()  throws IOException {
	   		
		    try
		    {
		      FileInputStream inputStream = new FileInputStream(excelFilePath);	      
		      Workbook workbook = new XSSFWorkbook(inputStream);
		      TestExecutionDetails = workbook.getSheetAt(1); 
		      for (int rowno = 0; rowno < TestExecutionDetails.getLastRowNum(); rowno++) {
		          Row nextRow = TestExecutionDetails.getRow(rowno + 1);
		          Cell TestScenarioID = nextRow.getCell(0);
		          int intTestScenarioID = (int)TestScenarioID.getNumericCellValue();
		          Cell packageName = nextRow.getCell(1);
		          String strpackageName = packageName.getStringCellValue();
		          Cell className = nextRow.getCell(2);
		          String strclassName = className.getStringCellValue();
		          Cell testScenarioName = nextRow.getCell(3);
		          String strtestScenarioName = testScenarioName.getStringCellValue();
		          Cell executeFlag = nextRow.getCell(4);
		          String strexecuteFlag = executeFlag.getStringCellValue();
		          Cell module = nextRow.getCell(5);
		          String strModule = module.getStringCellValue();      
		          lstTestExecutionsheet = new ArrayList();
		          lstTestExecutionsheet.add(strpackageName);
		          lstTestExecutionsheet.add(strclassName);
		          lstTestExecutionsheet.add(strtestScenarioName);
		          lstTestExecutionsheet.add(strexecuteFlag);
		          lstTestExecutionsheet.add(strModule);
		          mapTestExecution.put(Integer.valueOf(intTestScenarioID), lstTestExecutionsheet);
		       }
		        workbook.close();
		        inputStream.close();
		    } catch (Exception readTestExecutionException) {
		      System.out.println("readTestExecution Method Exception Message: " + readTestExecutionException);
		      throw new SkipException("Skipping execution - Issue with Test Execution Sheet");
			}
	  }
	  
	  /**
		This method is used to generate and run testNG xml file during the run time 
		@param suiteName - type of testsuite fetched from excel sheet
		@param testClassName - class name which consists test  to be executed
		@param testPath - path where test to be executed exists
		@param methodName - test to be executed
		*/
	  public void runTestNGSuite(String suiteName, String testClassName, String testPath, String methodName)
      {
        	
	      Map<String, String> testngParams = new HashMap();
	
	      TestNG myTestNG = new TestNG();
	      
	      XmlSuite mySuite = new XmlSuite();
	      mySuite.setName(suiteName);
	      
	      XmlTest myTest = new XmlTest(mySuite);
	      myTest.setName(testClassName);
	      myTest.setParameters(testngParams);
	      
	      List<XmlClass> myClasses = new ArrayList();
	      XmlClass xmlclass = new XmlClass(testPath);  
	      
	      List<XmlInclude> includedMethods = new ArrayList();
	      includedMethods.add(new XmlInclude(methodName));
	      xmlclass.setIncludedMethods(includedMethods);          
	      myClasses.add(xmlclass);         
	      myTest.setXmlClasses(myClasses);
	      
	      List<XmlTest> myTests = new ArrayList();
	      myTests.add(myTest);
	      mySuite.setTests(myTests);
	      
	      List<XmlSuite> mySuites = new ArrayList();
	      mySuites.add(mySuite);
	      myTestNG.setXmlSuites(mySuites);
	      myTestNG.run();
	      
	      myClasses.clear();
	      includedMethods.clear();
	      myTests.clear();
	      mySuites.clear();
      
    }
	  /**
		This method is used to read all the files and execute the test suite
		*/
	  public void testSuiteExecutionFlow() throws Exception{
		  
		  String strPackageName = "" ;
		  String strClassName="";
	      String strTestScenarioName ="";
	      String strExecuteFlag="";
	      String strModule="";	
	      
	    try
	    {
	      String[] arrConfigurationParameters = new String[14];
	      arrConfigurationParameters = readConfigurationProperties();      
	      String strBrowserName = arrConfigurationParameters[0].trim(); 
	      readEnvtDetails();
	      readTestData();
	      readObjectRepository();
	      readTestExecution();
	      PropertyConfigurator.configure(logPropertyPath);
	      String strEnvtName = lstEnvtsheet.get(0);
	      String strSuiteName = lstEnvtsheet.get(1);
	      String strProjectName = lstEnvtsheet.get(2);
	      strReportsFolderPath = createHTMLReportsFolder(strProjectName);
	      strExecutionDate = getCurrentDateTime("dd-MMM-yyyy");
	      for (int i = 1; i <= mapTestExecution.size(); i++) {
	        listkeyValue = (List)mapTestExecution.get(Integer.valueOf(i));
	         strPackageName = (String)listkeyValue.get(0);
	         strClassName = (String)listkeyValue.get(1);
	         strTestScenarioName = (String)listkeyValue.get(2);	         
	         strExecuteFlag = (String)listkeyValue.get(3);
	         strModule = (String)listkeyValue.get(4);		         
	        if (strExecuteFlag.equalsIgnoreCase("yes")) {
	          strResultStatus = "Pass";
	          String strExecutionStartTime = getCurrentDateTime("dd-MMM-yyyy") + " " + getCurrentDateTime("HH:mm:ss");
	          runTestNGSuite(strSuiteName, strClassName, strPackageName+"."+ strClassName, strTestScenarioName);
	          String strExecutionEndTime = getCurrentDateTime("dd-MMM-yyyy") + " " + getCurrentDateTime("HH:mm:ss");
	          String strTotalExecutionTime = getDateDifference(strExecutionStartTime, strExecutionEndTime, "dd-MMM-yyyy HH:mm:ss");
	          String strReportPath = generateHTMLReport(strProjectName, strEnvtName,strBrowserName, strClassName, strTestScenarioName, strReportsFolderPath, strExecutionStartTime, strExecutionEndTime, strTotalExecutionTime, sbDetailedReport);
	        }	        
	        listkeyValue.clear();
	      }	
	      generateSummaryReport(strProjectName, strReportsFolderPath, strBrowserName, strEnvtName);
	    }	    
	    catch (Exception testSuiteExecutionFlowException)
	    {
	      System.out.println("testSuiteExecutionFlow Method Exception Message: " + testSuiteExecutionFlowException);
	      throw new SkipException("Skipping execution - Test Suite Execution Failed");
	    }
	  }  
	  
	  
	 //******************** Object Repository methods ************
	   
	  /**
		This method is used to read the Object Repository excel sheet
		*/
	   public void readObjectRepository() throws IOException		     
		{
		    try
		    {	      
		      FileInputStream inputStream = new FileInputStream(new File(objectRepositoryFilePath));
		      Workbook workbook = new XSSFWorkbook(inputStream);		      	      	      
		      Sheet Sheets = workbook.getSheetAt(0);
		      for (int rowno = 0; rowno < Sheets.getLastRowNum(); rowno++) {
		          Row nextRow = Sheets.getRow(rowno + 1);	          
		          Cell objName = nextRow.getCell(0);
		          String strObjName = objName.getStringCellValue();
		          Cell objLocatorType = nextRow.getCell(1);
		          String strLocatorType = objLocatorType.getStringCellValue();
		          Cell objValue = nextRow.getCell(2);
		          String strValue = objValue.getStringCellValue();		          
		          lstLocators = new ArrayList();
		          lstLocators.add(strLocatorType);
		          lstLocators.add(strValue);
		          mapObjectRepository.put(strObjName, lstLocators);	          	          
		        }	      
		      workbook.close();
		      inputStream.close();
		    } catch (Exception readObjectRepositoryException) {
		      System.out.println("readObjectRepository Method Exception Message: " + readObjectRepositoryException);
	      throw new SkipException("Skipping execution - Issue with Object Repository Sheet");
			 }
		}
	   
	   /**
		This method is used to get the element locator 
		@param RuntimeLocator - value to be fetched at runtime else give as null("") 
		@param objectName - object name given in excel object repository
		@return ElementLocator - object locator
		*/
	   public By getElementLocator(String objectName,String RuntimeLocator) throws IOException	    
		{
		    List<String> getLocator = (List)mapObjectRepository.get(objectName);
		    String strLocatorType = (String)getLocator.get(0);
		    String strLocatorValue = null;	    
		    switch (strLocatorType.toLowerCase())
		    {
			    case "id": 
			    	ElementLocator = By.id((String)getLocator.get(1));
		    	break;
			    case "name": 
			    	ElementLocator = By.name((String)getLocator.get(1));
			    	break;
			    case "tagname":  
			    	ElementLocator = By.tagName((String)getLocator.get(1));
			    	break; 
			    case "classname": 
			    	 ElementLocator = By.className((String)getLocator.get(1));
			    	 break;
			    case "partiallinktext":  
			    	ElementLocator = By.partialLinkText((String)getLocator.get(1));
			    	break; 	    
			    case "cssselector":  
			    	ElementLocator = By.cssSelector((String)getLocator.get(1));
			    	break;
			    case "linktext":
			    	ElementLocator = By.linkText((String)getLocator.get(1));
			    	break; 
			    case "xpath":
			    	ElementLocator = By.xpath((String)getLocator.get(1));
			    	break;
			    case "cssselector+data": 
			    	if(RuntimeLocator.equalsIgnoreCase("")== false) {
			    		strLocatorValue = (String)getLocator.get(1);
			    		ElementLocator = By.cssSelector(appendObjectLocatorText(strLocatorValue,RuntimeLocator));
			    		break;
			    	}
			    	else {
				    	strLocatorValue = (String)getLocator.get(1);
				        ElementLocator = By.cssSelector(appendObjectLocatorText(strLocatorValue,""));
				        break; 
			    	}    
			    case "linktext+data":  
			    	if(RuntimeLocator.equalsIgnoreCase("")== false) {
			    		strLocatorValue = (String)getLocator.get(1);
			    		ElementLocator = By.linkText(appendObjectLocatorText(strLocatorValue,RuntimeLocator));
			    		break;
			    	}
			    	else {
			    		strLocatorValue = (String)getLocator.get(1);
				        ElementLocator = By.linkText(appendObjectLocatorText(strLocatorValue,""));
				        break; 
			    	}			    	
			    case "classname+data":
			    	if(RuntimeLocator.equalsIgnoreCase("")== false) {
			    		 strLocatorValue = (String)getLocator.get(1);
			    		 ElementLocator = By.className(appendObjectLocatorText(strLocatorValue,RuntimeLocator));
			    		break;
			    	}
			    	else {
				    	 strLocatorValue = (String)getLocator.get(1);
					     ElementLocator = By.className(appendObjectLocatorText(strLocatorValue,""));
					     break;
			    	}
			    case "partiallinktext+data":
			    	if(RuntimeLocator.equalsIgnoreCase("")== false) {
			    		strLocatorValue = (String)getLocator.get(1);
			    		ElementLocator = By.partialLinkText(appendObjectLocatorText(strLocatorValue,RuntimeLocator));
			    		break;
			    	}
			    	else {
				    	strLocatorValue = (String)getLocator.get(1);
				    	ElementLocator = By.partialLinkText(appendObjectLocatorText(strLocatorValue,""));
				    	break; 
			    	}
			    case "name+data":
			    	if(RuntimeLocator.equalsIgnoreCase("")== false) {
			    		strLocatorValue = (String)getLocator.get(1);
			    		ElementLocator = By.name(appendObjectLocatorText(strLocatorValue,RuntimeLocator));
			    		break;
			    	}
			    	else {
				    	strLocatorValue = (String)getLocator.get(1);
				        ElementLocator = By.name(appendObjectLocatorText(strLocatorValue,""));
				        break; 
			    	} 
			    case "id+data":
			    	if(RuntimeLocator.equalsIgnoreCase("")== false) {
			    		strLocatorValue = (String)getLocator.get(1);
			    		ElementLocator = By.id(appendObjectLocatorText(strLocatorValue,RuntimeLocator));
			    		break;
			    	}
			    	else {
				    	 strLocatorValue = (String)getLocator.get(1);
					     ElementLocator = By.id(appendObjectLocatorText(strLocatorValue,""));
					     break;
			    	}
			    case "xpath+data":
			    	if(RuntimeLocator.equalsIgnoreCase("")== false) {
			    		strLocatorValue = (String)getLocator.get(1);
			    		ElementLocator = By.xpath(appendObjectLocatorText(strLocatorValue,RuntimeLocator));
			    		break;
			    	}
			    	else {
				    	strLocatorValue = (String)getLocator.get(1);
				        ElementLocator = By.xpath(appendObjectLocatorText(strLocatorValue,""));
				        break; 
			    	}    
			    case "tagname+data":
			    	if(RuntimeLocator.equalsIgnoreCase("")== false) {
			    		strLocatorValue = (String)getLocator.get(1);
			    		ElementLocator = By.tagName(appendObjectLocatorText(strLocatorValue,RuntimeLocator));
			    		break;
			    	}
			    	else {
				    	strLocatorValue = (String)getLocator.get(1);
				        ElementLocator = By.tagName(appendObjectLocatorText(strLocatorValue,""));
				        break;
			    	}
				   }
		   return ElementLocator;
		 }
		
	   /**
		This method is used to append the locator with the dynamic text 
		@param strValue - dynamic text provided in excel sheet
		@param RuntimeLocator - value to be fetched at runtime else give as null("") 
		@return strNewLocatorValue - locator with appended text
		*/
	   public String appendObjectLocatorText(String strValue,String RuntimeLocator) throws IOException   
	    {
		   try {
			   String[] strSplittedValues = strValue.split("%");    
			   String strNewLocatorValue = "";
			   for (int i = 0; i < strSplittedValues.length; i++) {
				   if (i % 2 == 0) {
					   strNewLocatorValue = strNewLocatorValue + strSplittedValues[i];
				   }
				   else if (RuntimeLocator.equalsIgnoreCase("")== false) {
					   strNewLocatorValue = strNewLocatorValue + RuntimeLocator;
				   }
				   else if (mapTestData.containsKey(strSplittedValues[i])) {
					   strNewLocatorValue = strNewLocatorValue + (String)mapTestData.get(strSplittedValues[i]);
				   }
			   else {
				   strNewLocatorValue = strNewLocatorValue + strSplittedValues[i];
			   		}
		   }
		   return strNewLocatorValue;
		  }
		  catch (Exception e) {
			   System.out.println("Finding Locator Exception Message: " + e);
			   throw new SkipException("Skipping execution - unable to find test data reference for locator");
		  }	      
	    }
	    
	 //******************** Excel Handling methods ************
	   
	   /**
		This method is used to read the test data sheet 
		*/
	   public void readTestData() throws IOException {
			    
	    try
		    {	      
		      FileInputStream inputStream = new FileInputStream(new File(excelFilePath));		      
		      Workbook workbook = new XSSFWorkbook(inputStream);		       
			  Sheet Sheets = workbook.getSheetAt(2);		        
			  for (int rowno = 0; rowno < Sheets.getLastRowNum(); rowno++) {
	            Row nextRow = Sheets.getRow(rowno + 1);            
	            Cell ObjTestDataName = nextRow.getCell(0);
	            String strTestDataName = ObjTestDataName.getStringCellValue();
	            Cell ObjTestDataValue = nextRow.getCell(1);
	            String strTestDataValue = "";
	            switch (ObjTestDataValue.getCellType()) {
			        case 1: 
			          strTestDataValue = ObjTestDataValue.getStringCellValue();
			          break;
			        case 0: 
			        	if(HSSFDateUtil.isCellDateFormatted(ObjTestDataValue)) {
			        		Date strTestDataVal = ObjTestDataValue.getDateCellValue();
			        		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");			        		 			        	        
			        		strTestDataValue = dateFormat.format(strTestDataVal);
			        	}
			        	else {
			        		double dTestDataValue = ObjTestDataValue.getNumericCellValue();
					        strTestDataValue = Double.toString(dTestDataValue);
			        	}			          
			          break;
			        default: 
			          strTestDataValue = ObjTestDataValue.getStringCellValue();
	            }
	        mapTestData.put(strTestDataName, strTestDataValue);			
		  }
	      workbook.close();
	      inputStream.close();
	    } catch (NullPointerException noTestData) {
	      System.out.println("No TestData present : " + noTestData);
	     
	    }
	    catch (Exception readTestDataException) {
		      System.out.println("readTestData Method Exception Message: " + readTestDataException);
		      throw new SkipException("Skipping execution - Issue with Test Data Sheet");
		    }
	 }
	   
	   /**
		This method is used to get the value for the given test data from testdata excel sheet 
		@param strTestDataName - test data name provided in test data for which value needs to be fetched
		@return strTestDataValue - value stored for the given test data name in testdata sheet
		*/
	   public String getTestData(String strTestDataName)
	   	{
		   String strTestDataValue = (String)mapTestData.get(strTestDataName);
		   String strNewTestDataValue = "";
		   if (strTestDataValue != null) {
			   if (strTestDataValue.contains("%%")) {
				   String[] strsplitarray = strTestDataValue.split("%%");
			       for (int i = 0; i < strsplitarray.length; i++)
			         {
			           if (mapTestData.containsKey(strsplitarray[i])) {
			        	   strNewTestDataValue = strNewTestDataValue + (String)mapTestData.get(strsplitarray[i]);
			           } else {
			        	   strNewTestDataValue = strNewTestDataValue + strsplitarray[i];
			           }
			         }
			       strTestDataValue = strNewTestDataValue;
			   	}
	       return strTestDataValue;
		   }	     
	     return strTestDataName;
	   }
	
	   		  
	//******************** load properties file ************
	   
	   /**
		This method is used to load the configuration properties file
		*/
	   public void loadPropertiesFile() {
		   
		   try{
				 prop = new Properties();
				 FileInputStream fileip;
				 fileip = new FileInputStream(propertyFilePath) ;
				 prop.load(fileip); 
			}
			catch (FileNotFoundException e) {
			 	 e.printStackTrace();
				 throw new RuntimeException("Configuration properties file not found at " + propertyFilePath);
		 	 }	
			 catch (IOException e) {
					 e.printStackTrace();
			 } 
		   
	   }

	
	//******************** HTML Reporting Methods ************
	   
	   /**
		This method is used to get the current date and time
		@param format  - format in which date should be returned
		@return current date and time
		*/
	   public String getCurrentDateTime(String format)
	    {
		      Calendar cal = Calendar.getInstance();
		      SimpleDateFormat sdf = new SimpleDateFormat(format);
		      return sdf.format(cal.getTime());
	    }
		
	
	   /**
		This method is used to get the difference between two dates
		@param strDate1 - date 
		@param strDate2 - date
		@param strDateFormat - date format in which date should be returned
		@return strDateDiff - difference between two dates
		*/	
	   public String getDateDifference(String strDate1, String strDate2, String strDateFormat) throws Exception			      
	   		{
		      String strDateDiff = null;
		      try {
		        DateFormat format = new SimpleDateFormat(strDateFormat);
		        Date date1 = format.parse(strDate1);
		        Date date2 = format.parse(strDate2);
		        long diff = date2.getTime() - date1.getTime();
		        long diffSeconds = diff / 1000L % 60L;
		        long diffMinutes = diff / 60000L % 60L;
		        long diffHours = diff / 3600000L;
		        strDateDiff = diffHours + ":" + diffMinutes + ":" + diffSeconds;
		      } catch (Exception getDateDifferenceException) {
		        System.out.println("getDateDifference Method Exception Message: " + getDateDifferenceException);
			  }
			    return strDateDiff;
	   		}
	
	   /**
		This method is used to create Reports Folder if it does not exist
		@param strProjectName - project or release which is getting executed
		@return strPath - path where reports get generated
		*/
	   public String createHTMLReportsFolder(String strProjectName) {
			
		   String strPath = "./Results/Reports/"+ strProjectName+"_"+getCurrentDateTime("ddMMMyyyy") + "_" + getCurrentDateTime("HHmmss");
		    if (!new File(strPath).isDirectory())
		      new File(strPath).mkdirs();		     
		    return strPath;
			    
	   }
	   
	   /**
		This method is used to capture the screenshot for the failed test cases
		@param strTestName - test name which has been failed
		@return strScreenshotpath - provides the screenshot path
		*/
	   public String captureScreenshot(String strTestName)
	   	{
		    String strpath = "";
		    File directory = new File("./Results/Reports");
		    File[] fList = directory.listFiles();
		    File max = null;
		    for (File latestfile : fList) {
		    	if (latestfile.isDirectory() && (max == null || max.lastModified() < latestfile.lastModified())) {
		            max = latestfile;
		    	}
		    }
		    String strFileName = max.getName();
		    String screenshotFolderPath = "./Results/Reports/"+strFileName + "/TestScenarios/Screenshots";
		    try {
			      if (!new File(screenshotFolderPath).isDirectory()) {
			        new File(screenshotFolderPath).mkdirs();
			      }
			      String strcurrentdate = getCurrentDateTime("ddMMMyyyy") + "_" + getCurrentDateTime("HHmmss");
			      strpath = 
			    		  screenshotFolderPath + "/"+ strTestName + "_" + strcurrentdate + ".png";
			      File screenshot = (File)((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
			      File file = new File(strpath);
			      String strScreenshotName = strTestName + "_" + strcurrentdate + ".png";
			      String strScreenshotpath = "./Screenshots/" + strScreenshotName;
			      if (!file.exists())
			      {
			        FileUtils.copyFile(screenshot, file);
			      }
			      else {
			        System.out.println("Screenshot name already exists");
			        FileUtils.copyFile(screenshot, file);
			      }	      
			      return strScreenshotpath;
		    } 
		    catch (Exception captureScreenshotMethodException) {
		    	captureScreenshotMethodException.printStackTrace();
				throw new RuntimeException("Error occured while taking screenshot");
		    }
	   	}
	
	   /**
		This method is used to log the result(pass or fail) statements in report
		@param TestCaseName - test name (action which will be performed)
		@param ExpectedResult - expected behavior of the object
		@param ActualResult - actual behavior of the object
		@param Result - pass/fail status of the action performed
		@param strScreenshotPath - screenshot path for the failed test cases		
		*/
	   public void logResult(String TestCaseName, String ExpectedResult, String ActualResult, String Result, String strScreenshotPath)
		    throws IOException
	   	{		    
		    sbDetailedReport.append(" <tr>              <td>" + 
		      intSNo + "</td>" + 
		      "              <td>" + TestCaseName + "</td>" + 
		      "              <td>" + ExpectedResult + "</td>" + 
		      "              <td>" + ActualResult + "</td>");
		    
		
		    if (Result.equalsIgnoreCase("Pass")) {
		    	passStepsCount++;
		      sbDetailedReport.append("     <td id ='Status' style='background-color:#43B02A; color:White;'>" + Result + "</td>" + 
		        "          </tr>");
		    }
		    else {
		      strResultStatus = "Fail";
		      if (strScreenshotPath.equalsIgnoreCase("")) {
		    	  failStepsCount++;
		        sbDetailedReport.append("     <td id ='Status' style='background-color:Red; color:White;'>" + Result + "</td>" + 
		          "          </tr>");
		      }
		      else {
		    	  failStepsCount++;
		        sbDetailedReport.append("     <td id ='Status' style='background-color:Red;'><a style='color:White;' href='" + strScreenshotPath + "'><u>" + Result + "</u></a></td>" + 
		          "          </tr>");
		      }
		    }
		    
		    intSNo += 1;
	   	}
	
	   /**
		This method is used to generate HTML Report for each test executed
		@param ProjectName - project or release which is executed (fetched from testdata sheet)
		@param Environment - environment where test is executed
		@param Browser - browser where test is executed
		@param Functionality - functionality to which test belongs
		@param TestScenarioName - test name which is getting executed
		@param strReportPath - path where report gets generated
		@param strExecutionStartTime - execution start time
		@param strExecutionEndTime - execution end time
		@param strTotalExecutionTime - total execution time of test
		@param sbTable - detailed report steps of each test case
		@return	resultPath - report name along with path	
		*/
	   public String generateHTMLReport(String ProjectName, String Environment, String Browser, String Functionality, String TestScenarioName, String strReportPath, String strExecutionStartTime, String strExecutionEndTime, String strTotalExecutionTime, StringBuffer sbTable)
		    throws IOException
	     {
		    String resultFolderPath = strReportPath + "/TestScenarios";
		    String resultPath = resultFolderPath + "/" + getCurrentDateTime("ddMMMyyyy") + "_" + getCurrentDateTime("HHmmss") + "_" + TestScenarioName + ".html"; 		    		   
		    if (!new File(resultFolderPath).isDirectory())
			      new File(resultFolderPath).mkdirs();
		    int year = Calendar.getInstance().get(1);
		    StringBuffer sbReport = new StringBuffer();
		    if (strResultStatus.equalsIgnoreCase("Pass")) {
		      sbReport.append("<td id='overalltestresult' style='background-color:#43B02A; color:White; font-weight:bold;'>" + strResultStatus + "</td>");		    
		    } else {
		      sbReport.append("<td id='overalltestresult' style='background-color:Red; color:White; font-weight:bold;'>" + strResultStatus + "</td>");
		    }
		    int stepsCount = intSNo - 1;
		    int passCount = passStepsCount;
		    int failCount = failStepsCount;
		    File file = new File(resultPath);
		    if (!file.exists()) {
		      file.createNewFile();		      
		      FileWriter fout = new FileWriter(resultPath, true);
		      fout.write(
		        "<!DOCTYPE html><html lang='en'><head>  <META http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\">  <title>Detailed Execution Report</title><meta name='viewport' content='width=device-width, initial-scale=1'><link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/css/bootstrap.min.css'>"        
		        		+"  </head>" + 
		 		        "  <body style='padding:5px;'>" + 
		 		        "   <div class='container-fluid'>" + 
		 		        "    <div id='main'>" + 
		 		        "     <div style='float: left'>" + 
		 		        "       <img src='data:image/png;base64,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' alt='Logo' style='width:250px;height:55px;'>" + 
		 		        "     </div>" + 
		 		        "     <br/>" + 
		 		        "     <h2 style='border-bottom: 3px solid #86BC25;'>" +ProjectName+ "</h2>" + 		 		         
		 		        "     <br/>" +  		        		
			        "    <div>" + 
			        "\t\t\t\t\t\t\t<table id = 'SummaryTable' class='table table-bordered table-responsive' style='border:1.5px solid #86BC25' border='1000px'>" + 
			        "\t\t\t\t\t\t\t<tbody>" + 
			        "\t\t\t\t\t\t\t<tr>" + 
			        "\t\t\t\t\t\t\t\t<td style='font-weight:bold'>TestScenario Name</td>" + 
			        "\t\t\t\t\t\t\t\t<td id='TestScenarioName'>" + TestScenarioName + "</td>" + 
			        "\t\t\t\t\t\t\t\t<td style='font-weight:bold'>Functionality</td>" + 
			        "\t\t\t\t\t\t\t\t<td>" + Functionality + "</td>" + 
			        "\t\t\t\t\t\t\t</tr>" + 
			        "\t\t\t\t\t\t\t<tr>" + 
			        "\t\t\t\t\t\t\t\t<td style='font-weight:bold'>Environment</td>" + 
			        "\t\t\t\t\t\t\t\t<td>" + Environment + "</td>" + 
			        "\t\t\t\t\t\t\t\t<td style='font-weight:bold'>Browser</td>" + 
			        "\t\t\t\t\t\t\t\t<td>" + Browser + "</td>" + 			        
			        "\t\t\t\t\t\t\t</tr>" + 
			        "\t\t\t\t\t\t\t<tr>" + 
			        "\t\t\t\t\t\t\t\t<td style='font-weight:bold'>Execution Start Time</td>" + 
			        "\t\t\t\t\t\t\t\t<td id='StartTime'>" + strExecutionStartTime + "</td>" + 
			        "\t\t\t\t\t\t\t\t<td style='font-weight:bold'>Execution End Time</td>" + 
			        "\t\t\t\t\t\t\t\t<td id='EndTime'>" + strExecutionEndTime + "</td>" + 
			        "\t\t\t\t\t\t\t</tr>" + 
			        "\t\t\t\t\t\t\t<tr>" + 
			        "\t\t\t\t\t\t\t\t<td style='font-weight:bold'>Total Execution Time (HH:MM:SS)</td>" + 
			        "\t\t\t\t\t\t\t\t<td id='ExecutionTime'>" + strTotalExecutionTime + "</td>" + 
			        "\t\t\t\t\t\t\t\t<td style='font-weight:bold'>Test Result Status</td>" + 
			        sbReport + 
			        "\t\t\t\t\t\t\t</tr>" + 
			        "\t\t\t\t\t\t\t</tbody>" + 
			        "\t\t\t\t\t\t\t</table>" + 
			        "                          </div>" + 
			        "\t\t\t\t\t\t\t<table id = 'ReportTable' class='table table-bordered table-responsive' width='100%'>" + 
			        "\t\t\t\t\t\t\t<thead style='background-color:#D0D0CE'>" + 
			        "                          <tr>" + 
			        "                                          <th>SNo</th>" + 
			        "                                          <th>TestCase Name</th>" + 
			        "                                          <th>Expected Result</th>" + 
			        "                                          <th>Actual Result</th>" + 
			        "                                          <th>Status</th>" + 
			        "                          </tr>" + 
			        "\t\t\t\t\t\t\t</thead>" + 
			        sbTable + 
			        "\t\t\t\t\t\t\t</table>" + 
			        "\t\t\t\t\t\t    <footer id='foot01'></footer>" + 
			        "                          <div>" + 
			        "\t\t\t\t\t\t\t<table id = 'FooterTable' class='table table-bordered table-responsive' style='border:1.5px solid #86BC25' border='1000px'>" + 
			        "\t\t\t\t\t\t\t<tbody>" +
			        "\t\t\t\t\t\t\t<tr>" + 
			        "\t\t\t\t\t\t\t\t<td style='font-weight:bold'>Total Executed Steps :</td>" + 
			        "\t\t\t\t\t\t\t\t<td id='TotalSteps'>" + stepsCount + "</td>" + 
			        "\t\t\t\t\t\t\t\t<td style='font-weight:bold'>TestSteps Passed :</td>" + 
			        "\t\t\t\t\t\t\t\t<td id='PassStepsCount'>" + passCount + "</td>" +
			        "\t\t\t\t\t\t\t\t<td style='font-weight:bold'>TestSteps Failed :</td>" + 
			        "\t\t\t\t\t\t\t\t<td id='FailStepsCount'>" + failCount + "</td>" + 
			        "\t\t\t\t\t\t\t</tr>" + 
			        "\t\t\t\t\t\t\t</tbody>" + 
			        "\t\t\t\t\t\t\t</table>" + 
			        "                          </div>" + 
			        "\t\t\t\t\t\t\t</body>" + 
			        "\t\t\t\t\t\t\t</html>");
			      
			      fout.close();
			      sbTable.delete(0, sbTable.length());
			      sbDetailedReport.delete(0, sbDetailedReport.length());
			    intSNo =1;
			    passStepsCount =0;
			    failStepsCount = 0;
			    }
			    else {
			      System.out.println("Detailed report file already exists");
			    }
			    return resultPath;
		  	}
	
	
	   /**
		This method is used to calculate summary of all tests executed(pass/ fail count)
		@param ProjectName - project or release which is executed (fetched from testdata sheet)
		@param strReportPath - path where report gets generated
		@param Browser - browser where test is executed
		@param Environment - environment where test is executed		
		*/
	   public void generateSummaryReport(String ProjectName, String strReportPath, String strBrowser, String strEnvironment)
		    throws Exception
		  {
		    String resultFolderPath = strReportPath + "/TestScenarios";
		    File directory = new File(resultFolderPath);
		    File[] fList = directory.listFiles();
		    int intPassCount = 0;
		    int intFailCount = 0;
		    int intExecutionCount = 0;
		    int intTotalStepsCount  = 0;
		    Document doc = null;
		    String flags = null;
		    StringBuffer sbHTML = new StringBuffer();
		    for (File file : fList) {
		      String strFileName = file.getName();
		      if(strFileName.equalsIgnoreCase("Screenshots")==false) {		     
			      doc = Jsoup.parse(file, "UTF-8");
			      Elements elementsStatus = doc.select("td#Status");
			      String DetailReportfilePath = "./TestScenarios/" + strFileName;
			      String strStartTime = doc.select("td#StartTime").first().text();
			      String strEndTime = doc.select("td#EndTime").first().text();
			      String strExecTime = doc.select("td#ExecutionTime").first().text();
			      String strTestScenarioName = doc.select("td#TestScenarioName").first().text();
			      String strTotalStepsCount = doc.select("td#TotalSteps").first().text();
			      intTotalStepsCount = Integer.parseInt(strTotalStepsCount) + intTotalStepsCount;	
			      for (Element eleStatus : elementsStatus) {
			        String strStatus = eleStatus.select("td#Status").first().text();
			        if (strStatus.equalsIgnoreCase("pass")) {
			          flags = "Pass";
			        } 
			        else {
			        	flags = "Fail";
				          break;
			        }		          
			      }
			      if (flags == "Pass") {
			        intPassCount++;
			      } else if (flags == "Fail") {
			        intFailCount++;
			      } 		      
			      StringBuffer statusbuffer = new StringBuffer();
			      if (flags.equalsIgnoreCase("Pass")) {
			        statusbuffer.append("<td style='background-color:#43B02A; color:White;'><a href='" + DetailReportfilePath + "'>" + flags + "</a></td>");
			      }  else {
			        statusbuffer.append("<td style='background-color:Red; color:White;'><a href='" + DetailReportfilePath + "'>" + flags + "</a></td>");
			      }
			      sbHTML.append("    <tr>      <td>"+Sno+"</td>      "+
			      	"      <td><a href='" + DetailReportfilePath + "'><font color='blue'>" + strTestScenarioName + "</font></a></td>" + 
			      	"      <td>" + strStartTime + "</td>" + 
			        "      <td>" + strEndTime + "</td>" + 
			        "      <td>" + strExecTime + "</td>" + 
			        statusbuffer + 
			        "   </tr>");		      
			      sumExecutionTime(strExecTime);
			      Sno += 1;
		      }
		    }
		    intExecutionCount = intPassCount + intFailCount ;		    
		    int[] arrStatus = { intPassCount, intFailCount, intExecutionCount , intTotalStepsCount };		    
		    if (intExecutionCount > 0) {
		      intTcsExecutedCount = intExecutionCount;
		      if (intFailCount > 0) {
		        strExecutionStatus = "FAIL";
		      } else {
		        strExecutionStatus = "PASS";
		      }
		    }		    
		    createSummaryReport(ProjectName, strBrowser, strEnvironment, strReportPath, sbHTML, arrStatus);
		  }
	
	   /**
		This method is used to create summary of all tests executed
		@param ProjectName - project or release which is executed (fetched from testdata sheet)		
		@param Browser - browser where tests are executed
		@param Environment - environment where tests are executed	
		@param strReportPath - path where report gets generated
		@param masterTable - contains detailed steps of each test executed
		@param arrStatus - execution status	(PassCount, FailCount, ExecutionCount)
		*/
	   public void createSummaryReport(String ProjectName, String strBrowser, String Environment, String strReportPath, StringBuffer masterTable, int[] arrStatus)
		    throws IOException
		  {
		    String resultPath = strReportPath + "/" + getCurrentDateTime("ddMMMyyyy") + "_" + getCurrentDateTime("HHmmss") +"_"+ProjectName+"_"+"SummaryReport.html";
		    int year = Calendar.getInstance().get(1);
		    File file = new File(resultPath);
		    if (!file.exists()) {
		      file.createNewFile();
		      FileWriter fout = new FileWriter(resultPath, true);
		      fout.write("  <!DOCTYPE html> <html lang='en'>  <head>         <META http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\">         <title>Test Automation Summary Report</title>         <script type='text/javascript' src='https://www.gstatic.com/charts/loader.js'></script>\t\t  <style>#main{background-color:#FFF;border-radius:0 0 5px 5px}h1{font-family:Georgia,serif;border-bottom:3px solid #0F0;color:#960;font-size:30px}table{width:100%}table#'ReportTable',td#'ReportTable',th#'ReportTable'{border:1px solid lightblue;border-collapse:collapse;padding:5px}table#'SummaryTable',td#'SummaryTable',th#'SummaryTable'{style=border:10px solid black;background-color:white;border-collapse:collapse;padding:5px;border-style: solid;border-width: 2px 10px 4px 20px;}h4{width:400px}th{text-align:left}#SummaryTable tr:nth-child(even){background-color:#fff}#SummaryTable tr:nth-child(odd){background-color:#fff}#SummaryTable td:nth-child(odd){background-color:#fff}.pass{color:#fff;background-color:#43B02A}.fail{color:#fff;background-color:red} a{color:white !important;text-decoration: underline;}; a:link{color:white !important;text-decoration: underline;} ; a:visited{color:white !important; text-decoration: underline;}; a:hover{color:white !important;text-decoration: underline;}</style>         <script src='http://code.jquery.com/jquery-1.9.1.js'> </script>         <meta name='viewport' content='width=device-width, initial-scale=1'>         <link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/css/bootstrap.min.css'>         <script src='https://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/js/bootstrap.min.js'></script>         <script type='text/javascript'>google.charts.load('current', {'packages':['corechart']});google.charts.setOnLoadCallback(drawChart);function drawChart() {var data = new google.visualization.DataTable();data.addColumn('string', 'Status');data.addColumn('number', 'Count');data.addRow(['Pass', " + 
	
	
		        arrStatus[0] + "]);data.addRow(['Fail', " + arrStatus[1] + "]);var options = {'title':'Execution Summary',colors: ['#43B02A', 'red', '#FFD700']};var chart = new google.visualization.PieChart(document.getElementById('piechart'));chart.draw(data, options);}</script>" + 
		        "  </head>" + 
		        "  <body style='padding:5px;'>" + 
		        "   <div class='container-fluid'>" + 
		        "    <div id='main'>" + 
		        "     <div style='float: left'>" + 
		        "       <img src='data:image/png;base64,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' alt='Logo' style='width:250px;height:55px;'>" + 
		        "     </div>" + 
		        "     <br/>" + 
		        "     <h2 style='border-bottom: 3px solid #86BC25;'>Test Execution Summary Report</h2>" + 
		        "     <h3>" + ProjectName + "</h3>" + 
		        "     <br/>" + 
		        "     <table style='border:0'><tbody><tr><td align='top' style='vertical-align:top'>" + 
		        "      <div style='background-color:white;vertical-align:top;'>" + 
		        "       <table id = 'SummaryTable' class='table table-bordered table-responsive' style='border:1.5px solid #86BC25' border='1000px'>" + 
		        "        <tbody>" + 
		        "         <tr><td style='font-weight:bold'>Test Execution Environment </td> <td>" + Environment + "</td></tr>" +
		        "         <tr><td style='font-weight:bold'>Test Execution Browser </td> <td>" + strBrowser + "</td></tr>" + 		        
		        "         <tr><td style='font-weight:bold'>Executed TestScenarios Count </td> <td>" + arrStatus[2] + "</td></tr>" + 
		        "         <tr><td style='font-weight:bold'>Executed TestSteps Count </td> <td>" + arrStatus[3] + "</td></tr>" + 
		        "         <tr><td style='font-weight:bold'>TestScenarios Pass Count </td> <td>" + arrStatus[0] + "</td></tr>" + 
		        "         <tr><td style='font-weight:bold'>TestScenarios Fail Count </td> <td>" + arrStatus[1] + "</td></tr>" + 
		        "         <tr><td style='font-weight:bold'>Test Execution Date </td> <td>" + strExecutionDate + "</td></tr>" + 
		        "         <tr><td style='font-weight:bold'>Total Execution Time (HH:MM:SS)</td> <td>" + iHours + ":" + iMinutes + ":" + iSeconds + "</td></tr>" + 
		        "        </tbody>" + 
		        "       </table>" + 
		        "      </div>" + 
		        "      </td>" + 
		        "      <td align='right' style='vertical-align:top'>" + 
		        "       <div id='piechart' style='width: 600px; height: 310px;'></div>" + 
		        "      </td></tr></tbody></table>" + 
		        "      <table id = 'ReportTable' class='table table-bordered table-responsive'>" + 
		        "       <thead style='background-color:#D0D0CE'>" + 
		        "        <tr>" + 
		        "         <th>SNo</th>" + 
		        "         <th>TestScenario Name</th>" + 
		        "         <th>Execution Start Time</th>" + 
		        "         <th>Execution End Time</th>" + 
		        "         <th>Total Execution Time (HH:MM:SS)</th>" + 
		        "         <th>Status</th>" + 
		        "        </tr>" + 
		        "       </thead>" + masterTable + "</table>" + 
		        "      <footer id='foot01'></footer>" + 
		        "   </body>" + 
		        "  </html>");
		      
		      fout.close();		            
		    }
		    else {
		      System.out.println("Summary Report File already exists");
		    }
		  }
	// ******************** load properties file ************

			/**
			 * This method is used to load the configuration properties file
			 */
			public void loadPropertiesFiles() {

				try {
					prop = new Properties();
					InputStream fileip = getClass().getClassLoader().getResourceAsStream(propertyFilePath);
					/*
					 * FileInputStream fileip; fileip = new FileInputStream(propertyFilePath) ;
					 */
					prop.load(fileip);
					logged.info("Loaded properties file");
				} catch (FileNotFoundException e) {
					e.printStackTrace();
					throw new RuntimeException("Configuration properties file not found at " + propertyFilePath);
				} catch (IOException e) {
					e.printStackTrace();
				}

			}

			// ******************** Read all the values from excel and OR sheets
			// ************

			/**
			 * This method is used to read the configuration properties file and excel data
			 */
			public void fetchExecutionData() {

				try {
					logged.info("Read test data and object repository");
					readEnvtDetails();
					readTestData();
					readObjectRepository();
					readTestExecution();
					PropertyConfigurator.configure(logPropertyPath);
				} catch (Exception e) {
					System.out.println("testSuiteExecutionFlow Method Exception Message: " + e);
					throw new SkipException("Skipping execution - Test Suite Execution Failed");
				}

			}
	   /**
		This method is used to calculate sum of Execution time
		@param strTime - total execution time	
		*/
	   public void sumExecutionTime(String strTime)
			{
			    int irem = 0;
			    int iquo = 0;
			    if (strTime.trim().length() > 0) {
			      String[] arrStr = strTime.split(":");
				      iHours += Integer.parseInt(arrStr[0]);
				      iMinutes += Integer.parseInt(arrStr[1]);
				      iSeconds += Integer.parseInt(arrStr[2]);
				      if (iSeconds > 60) {
				        irem = iSeconds % 60;
				        iquo = iSeconds / 60;
				        iSeconds = irem;
				        iMinutes += iquo;
				      }
				      
				      if (iMinutes > 60) {
				        irem = iMinutes % 60;
				        iquo = iMinutes / 60;
				        iMinutes = irem;
				        iHours += iquo;
				      }
				    }
	    	}
	   	   
		}
	
