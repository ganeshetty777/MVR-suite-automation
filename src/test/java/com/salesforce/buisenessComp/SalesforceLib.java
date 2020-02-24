package com.salesforce.buisenessComp;

import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.firefox.FirefoxProfile;

import src.com.salesforce.genericLib.WebDriverCommonUtils;

import com.salesforce.genericLib.Driver;
import com.salesforce.genericLib.ExcelLib;



public class SalesforceLib {
	
	ExcelLib eLib=new ExcelLib();
	WebDriverCommonUtils wUtils=new WebDriverCommonUtils();
	public static final String INPUT_FILE="C:\\MVR_Report_Automation\\InputFile.xlsx";
	
	public  String url() throws InvalidFormatException, IOException{
    	
		String url=eLib.getexcelMasterMetaData("Login", 2, 2, SalesforceLib.INPUT_FILE);
		return url;
		
	}
   
   public String mvrReportPath() throws InvalidFormatException, IOException{
	   
	    String root="C:\\MVR_Report_Automation\\MVR Report\\";
		String excelpath1=eLib.getexcelMasterMetaData("Login", 2, 3, SalesforceLib.INPUT_FILE);
		
		return root+excelpath1+".xls";
		
	}
    
   public String outputReportPath() throws InvalidFormatException, IOException{
	   
		String root="C:\\MVR_Report_Automation\\Output Report\\";
		String excelpath2=eLib.getexcelMasterMetaData("Login", 2, 4, SalesforceLib.INPUT_FILE);
		
		return root+excelpath2+".xls";
		
	}
	
	
	public void login(String userName , String password) throws InvalidFormatException, IOException{
		
		System.setProperty("webdriver.chrome.driver", "./src/main/resources/chromedriver.exe");
		
		 ChromeOptions options = new ChromeOptions();  
	        options.addArguments("--browser.download.folderList=2");
	        options.addArguments("--browser.helperApps.neverAsk.saveToDisk=application/vnd.ms-excel");
	        options.addArguments("--browser.download.dir=D:\\Java\\MVR_Report");
	/*	FirefoxProfile profile=new FirefoxProfile();
	    profile.setPreference("browser.download.folderList", 2);
	    profile.setPreference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel");
	    profile.setPreference("browser.download.dir", "D:\\Java\\MVR_Report");*/
	    
	    /*FirefoxOptions options= new FirefoxOptions();
	    options.setProfile(profile);
	    
	    WebDriver driver= new FirefoxDriver(options);*/
	    
		 SalesforceLib sLib=new SalesforceLib();		
		 Driver.driver.manage().window().maximize();
		 Driver.driver.get(sLib.url());
		 Driver.driver.manage().window().maximize();
		 Driver.driver.findElement(By.xpath("//input[@id='username']")).sendKeys(userName);
		 Driver.driver.findElement(By.xpath("//input[@id='password']")).sendKeys(password);
		 Driver.driver.findElement(By.xpath("//input[@id='Login']")).click();
		 
		 wUtils.waitFortextPresent("//a[text()='Home']", "Home");
			
		 
		 }
	
    public int getCelldata(String SheetName ,int rowNum, int colNum,String excelpath) throws InvalidFormatException, IOException{
		
	String rNumber=eLib.getexcelMasterMetaData(SheetName, rowNum, colNum, excelpath);
	int data=Integer.parseInt(rNumber); 
		return data;
		
      }
    
    

}
