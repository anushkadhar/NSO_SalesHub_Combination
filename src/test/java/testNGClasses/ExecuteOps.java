package testNGClasses;

import testBase.BaseClass;
import testBase.CommonActions;

import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import pages.ReusableComponents;

public class ExecuteOps extends BaseClass{
public ExecuteOps(){}
CommonActions common = new CommonActions();
ReusableComponents comp=new ReusableComponents();

@BeforeTest
public void setBaseURL() throws Exception {
	fetchExecutionData();
	loadPropertiesFiles();
	common.invokeBrowser();
	common.setEnvtURL();
	logged.info("Execution started");
}

@Test
public void execute() throws Exception
{
	comp.createList();
	//comp.createDocumentLibrary();
	//comp.createDocument();
	//comp.uploadFunctionality();
}

@AfterTest
public void endSession() {
	common.closeBrowser();
	logged.info("Execution stopped");
}

}
