package src.com.salesforce.genericLib;

import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.By;
import org.openqa.selenium.Capabilities;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.support.events.EventFiringWebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.salesforce.buisenessComp.SalesforceLib;
import com.salesforce.genericLib.Driver;
import com.salesforce.genericLib.ExcelLib;

public class WebDriverCommonUtils {
	

	Date date = new Date();
	ExcelLib eLib=new ExcelLib();
	
	public void waitForElementPresent(String webElementXpath){
		 WebDriverWait wait = new WebDriverWait(Driver.driver, 4);
		 wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(webElementXpath)));
		 
		 }

	
	
	public void waitForElementClickable(String webElementXpath){
		 WebDriverWait wait = new WebDriverWait(Driver.driver, 4); 
		 wait.until(ExpectedConditions.elementToBeClickable(By.xpath(webElementXpath)));
		 
		 }
	
	 public void waitFortextPresent(String webElementXpath,String text){
		 WebDriverWait wait = new WebDriverWait(Driver.driver, 4); 
		 wait.until(ExpectedConditions.textToBePresentInElementLocated(By.xpath(webElementXpath), text));
		 
		 }
	
	
	public void implicitWait(){
		
		Driver.driver.manage().timeouts().implicitlyWait(3, TimeUnit.SECONDS);
	
	}
	
	/*public static Capabilities profile(){
		System.setProperty("webdriver.gecko.driver", "C:\\com.MVRAutomation.org\\src\\main\\resources\\geckodriver.exe");
		
		
		FirefoxProfile profile=new FirefoxProfile();
	    profile.setPreference("browser.download.folderList", 2);
	    profile.setPreference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel");
	    profile.setPreference("browser.download.dir", "D:\\Java\\MVR_Report");
	    
	    FirefoxOptions options= new FirefoxOptions();
	    options.setProfile(profile);
	    
	    WebDriver driver= new FirefoxDriver(options);
	    return (Capabilities) profile;
		
	}*/
	
	
	public void captureScreenshot(DateFormat dateFormat,int i) throws IOException, InvalidFormatException{
		
		 String presentationName=eLib.getexcelMasterMetaData("Login", 2, 5, SalesforceLib.INPUT_FILE);
	     EventFiringWebDriver eDriver=new EventFiringWebDriver(Driver.driver);
	     File srcImgFile=eDriver.getScreenshotAs(OutputType.FILE);
	     File dstImgPath=new File("C:\\MVR_Report_Automation\\Screenshots\\"+presentationName+"_"+dateFormat.format(date)+"\\"+"Slide_"+i+".jpeg");
	     FileUtils.copyFile(srcImgFile, dstImgPath); 
	}
	
   public static Set<String> findDuplicates(List<String> listContainingDuplicates) {
			 
			final Set<String> setToReturn = new HashSet<String>();
			final Set<String> set1 = new HashSet<String>();
	 
			for (String yourInt : listContainingDuplicates) {
				if (!set1.add(yourInt)) {
					setToReturn.add(yourInt);
				}
			}
			return setToReturn;
		}
		

}
