package src.com.salesforce.SingleRun;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.RowIdLifetime;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebElement;
import org.testng.Reporter;
import org.testng.SkipException;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.salesforce.buisenessComp.ReportStatus;
import com.salesforce.buisenessComp.SalesforceLib;
import com.salesforce.genericLib.Driver;
import com.salesforce.genericLib.ExcelLib;
import com.salesforce.genericLib.SendMail_PostmarkApp;

import src.com.salesforce.genericLib.WebDriverCommonUtils;



public class MVRTest_SingleRun_ClmRteDevEnv{
	
	
	SalesforceLib sLib=new SalesforceLib();
	ExcelLib eLib=new ExcelLib();
	WebDriverCommonUtils wUtils=new WebDriverCommonUtils();
	ReportStatus actStatus=new ReportStatus();
	List<WebElement> slideCount;
	int totalRowCountMvrs;
	static int count=0;
	static JFrame frmOpt;
	String sandbox;
	String status; 
	String subjectStatus;
	boolean flag = false;
	
	
	@BeforeTest
	public void migration() throws InterruptedException, Exception, IOException {
		
		String presentationNameDev = eLib.getexcelMasterMetaData("Login", 2, 5, SalesforceLib.INPUT_FILE);
		
		int condition1=0;
		
		condition1 = JOptionPane.showConfirmDialog(null, "Do you want to run Migration Suite for the build?", 
                presentationNameDev, JOptionPane.YES_NO_OPTION);
		if(condition1==0){
			
			Reporter.log(presentationNameDev+" iDetail is proceeding with Migration Suite", true);
			
			//Driver.driver.manage().timeouts().implicitlyWait(120, TimeUnit.SECONDS);
			
			Reporter.log("Browser launched with given URL for migration", true);
			
			String userNameDev = eLib.getexcelMasterMetaData("Login", 3, 0, SalesforceLib.INPUT_FILE);
			String passwordDev = eLib.getexcelMasterMetaData("Login", 3, 1, SalesforceLib.INPUT_FILE);
			
			String userNameTest = eLib.getexcelMasterMetaData("Login", 2, 0, SalesforceLib.INPUT_FILE);
			String passwordTest = eLib.getexcelMasterMetaData("Login", 2, 1, SalesforceLib.INPUT_FILE);
			
			sLib.login(userNameDev, passwordDev);
			
			Reporter.log("username, password entered", true);
			
			Driver.driver.findElement(By.id("sbstr")).sendKeys(presentationNameDev);
			
			Driver.driver.findElement(By.name("search")).click();
			
			Reporter.log(presentationNameDev+" iDetail is searched for the first time", true);
			
			Driver.driver.findElement(By.xpath("//div[@id='Clm_Presentation_vod__c_body']/table/tbody/tr[2]/th/a")).click();
			
			Thread.sleep(1000);
			
			Driver.driver.findElement(By.id("sbstr")).sendKeys(presentationNameDev);
			
			
			try{
				  
				  wUtils.waitForElementPresent("//a[span[img[@alt='CLM Presentation']]/following-sibling::strong[text()="+"'"+presentationNameDev+"'"+"]]");
				  String actPresName=Driver.driver.findElement(By.xpath("//a[span[img[@alt='CLM Presentation']]/following-sibling::strong[text()="+"'"+presentationNameDev+"'"+"]]")).getText();
				  
				  System.out.println("Presentation Name: "+actPresName.trim());
				  if(actPresName.trim().equals(presentationNameDev)){
					  
					wUtils.waitForElementClickable("//a[span[img[@alt='CLM Presentation']]/following-sibling::strong[text()="+"'"+presentationNameDev+"'"+"]]");
					Driver.driver.findElement(By.xpath("//a[span[img[@alt='CLM Presentation']]/following-sibling::strong[text()="+"'"+presentationNameDev+"'"+"]]")).click();
					  
					 Reporter.log("Presentation found and seleceted", true);
					  
					 
					
				  }
				  else{
					  throw new SkipException("Duplicate presentation found OR presentation doesn't exist");
				  }
				  
			  }
			  catch(Exception e){
				  e.printStackTrace();
				  Reporter.log("Duplicate presentation found or presentation doesn't exist", true);
				  if (frmOpt == null) {
				        frmOpt = new JFrame();
				    }
				    frmOpt.setVisible(true);
				    frmOpt.setLocationRelativeTo(null);
				    frmOpt.setAlwaysOnTop(true);
					JOptionPane.showMessageDialog(frmOpt, "Duplicate presentation found or presentation doesn't exist");
				    Driver.driver.quit();
				    throw new SkipException("Duplicate presentation found or presentation doesn't exist"); 
			  }
			
			
			////////////////////GUID Comparison//////////////////////
			
			String presentationMvrGuid = eLib.getexcelMasterMetaData("Presentation", sLib.getCelldata("MVRP", 1, 0, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRP", 1, 6, SalesforceLib.INPUT_FILE), sLib.mvrReportPath());
			
			String presentationIdDev = Driver.driver.findElement(By.xpath("//td[span[contains(text(),'Presentation Id')]]/following-sibling::td[1]/div")).getText();
			
			if (presentationMvrGuid.equals(presentationIdDev))
				
			{
				Reporter.log("Since Guid in MVR " +presentationMvrGuid+" is matching with Presentation GUID in SF "+presentationIdDev+ " proceeding with Migration ", true);
			}
			
			else
			{
				Driver.driver.close();
				Reporter.log("Since Guid in MVR " +presentationMvrGuid+" is mis-matching with with Presentation GUID in SF " +presentationIdDev+ " quitting the browser", true);
				throw new SkipException("Since Guid in MVR " +presentationMvrGuid+" is mis-matching with with Presentation GUID in SF " +presentationIdDev+ " quitting the browser");	
			}
			
            ////////////////////GUID Comparison//////////////////////
			
			Driver.driver.findElement(By.name("migrate_vod")).click();
			
			WebElement fr = Driver.driver.findElement(By.id("vod_iframe"));
			Driver.driver.switchTo().frame(fr);
			
			Driver.driver.findElement(By.xpath("//option[contains(text(), 'Sandbox')]")).click();
			
			Reporter.log(presentationNameDev+" iDetail Migration Initiated", true);
	        
			Driver.driver.findElement(By.name("username")).sendKeys(userNameTest);
			Driver.driver.findElement(By.name("password")).sendKeys(passwordTest);
			Driver.driver.findElement(By.id("submitButton")).click();
			
			try
			{
			    wUtils.waitFortextPresent("//div[@id='result_div']","Migration Error:");
			    
			    if(Driver.driver.findElement(By.xpath("//div[@id='result_div']")).isDisplayed());
			    {
			    	String msg = Driver.driver.findElement(By.xpath("//div[@id='result_div']")).getText();
					JOptionPane.showMessageDialog(null, msg);
					Reporter.log(presentationNameDev+" iDetail Migration Suite is not Proceeding Any Further.", true);
					Driver.driver.close();
			    }
			}
			
			catch(Exception e)
			{	
				
			}
			
			try
			{
				Driver.driver.findElement(By.xpath("//*[@id='submitButton' and @value='Continue']")).click();
			}
			catch(Exception e)
			{
				throw new SkipException(presentationNameDev+" iDetail Migration Suite is not Proceeding Any Further.");
			}
			
			boolean flag2=true;
			int condition = 0;
			
			try
			{
				Thread.sleep(40000);
				wUtils.waitFortextPresent("//div[@id='status_div']","Migration completed successfully");
				
				if (frmOpt == null) {
			        frmOpt = new JFrame();
			        JOptionPane.showMessageDialog(frmOpt, "Migration Completed Successfully");
			    }
				
				Reporter.log(presentationNameDev+" iDetail Migration is Completed and proceeding with Operational Analytical Suite", true);
				
				flag2=false;
				
				condition = JOptionPane.showConfirmDialog(null, "Do you want to run Operational Analytical suite for the build?", 
	                    presentationNameDev, JOptionPane.YES_NO_OPTION);
				if(condition==0){
					Thread.sleep(10000);
					Reporter.log(presentationNameDev+" iDetail is proceeding with Operational Analytical Suite", true);
						}
				else
				{
					Driver.driver.close();
					Reporter.log(presentationNameDev+" iDetail Operational Analytical Suite is not initiated.", true);
					throw new SkipException(presentationNameDev+" iDetail Operational Analytical Suite is not initiated.");	
				}
			}
			
			catch (Exception e)
			{
					
			}
			
			if (flag2==true)
			{
				flag = Driver.driver.findElement(By.xpath("//*[@id='submitButton' and @value='Overwrite']")).isDisplayed();
				if(flag==true){
					 condition = JOptionPane.showConfirmDialog(null, "Do you want to overwrite the build present in test env?", 
	                        presentationNameDev, JOptionPane.YES_NO_OPTION);
					if(condition==0){
						Driver.driver.findElement(By.xpath("//*[@id='submitButton' and @value='Overwrite']")).click();
						
						Thread.sleep(40000);
						
						wUtils.waitFortextPresent("//div[@id='status_div']","Migration completed successfully");
						
						if (frmOpt == null) {
					        frmOpt = new JFrame();
					        JOptionPane.showMessageDialog(frmOpt, "Migration Completed Successfully");
					    }
						
						Reporter.log(presentationNameDev+" iDetail Migration is Completed and proceeding with Operational Analytical Suite", true);
							}
					else
					{
						Reporter.log(presentationNameDev+" iDetail Migration is In-Complete and proceeding with Operational Analytical Suite", true);
					}
			}
					
					condition = JOptionPane.showConfirmDialog(null, "Do you want to run Operational Analytical suite for the build?", 
	                        presentationNameDev, JOptionPane.YES_NO_OPTION);
					if(condition==0){
						Reporter.log(presentationNameDev+" iDetail is proceeding with Operational Analytical Suite", true);
							}
					else
					{
						Driver.driver.close();
						Reporter.log(presentationNameDev+" iDetail Operational Analytical Suite is not initiated.", true);
						throw new SkipException(presentationNameDev+" iDetail Operational Analytical Suite is not initiated.");	
					}
				}
				}
		else
		{
			Reporter.log(presentationNameDev+" iDetail Migration is not being initiated and proceeding with Operational Analytical Suite", true);
		}
	} 

	@Test
    public void login() throws InvalidFormatException, IOException, InterruptedException {
		
		File file=new File("C:\\MVR_Report_Automation\\Screenshots");
		File file1=new File("C:\\MVR_Report_Automation\\Output Report");
		try{
		FileUtils.cleanDirectory(file);
		FileUtils.cleanDirectory(file1);
		}
        catch(IllegalArgumentException e){
			
			JOptionPane.showMessageDialog(null, "Before Execution please place all Subfolders inside C:\\MVR_Report_Automation");
			throw new SkipException("Before Execution please place all Subfolders inside C:\\MVR_Report_Automation");
		}
		catch(Exception e){
			
			JOptionPane.showMessageDialog(null, "Please close the MVR Output Report/Screenshots file");
			throw new SkipException("Please close the MVR Output Report/Screenshots file");
		}
		
		File srcFile=new File("Output_Report_Template.xls");
	    File destFile=new File(sLib.outputReportPath());
	    FileUtils.copyFile(srcFile, destFile);
	    
	    Reporter.log("OutPut Report template copied", true);
	    
	    File mvrFile=new File(sLib.mvrReportPath());
	    if(mvrFile.exists()){
	    
		File folder = new File("C:\\MVR_Report_Automation\\MVR Report");
		File[] listOfFiles = folder.listFiles();

		    for (int i = 0; i < listOfFiles.length; i++) {
		      if (listOfFiles[i].isFile()) {
		        System.out.println("File :" + listOfFiles[i].getName());
		        count++;
		        
		      } else if (listOfFiles[i].isDirectory()) {
		        System.out.println("Directory :" + listOfFiles[i].getName());
		        count++;
		      }
		    }System.out.println("Total file :"+count);
		    
		    if(count==1){
		    	
		  String userName=eLib.getexcelMasterMetaData("Login", 2, 0, SalesforceLib.INPUT_FILE);
		  String password=eLib.getexcelMasterMetaData("Login", 2, 1, SalesforceLib.INPUT_FILE);
		  String presentationName=eLib.getexcelMasterMetaData("Login", 2, 5, SalesforceLib.INPUT_FILE);
		  String presentationID=eLib.getexcelMasterMetaData("Login", 2, 9, SalesforceLib.INPUT_FILE);
		 
		  sLib.login(userName, password);
		  Reporter.log("Browser launched with given URL and username, password entered", true);
		  
		  
		  Driver.driver.findElement(By.xpath("//input[@autocomplete='off']")).sendKeys(presentationID);
		  
		  
		  /////////
		  
		  Driver.driver.findElement(By.name("search")).click();
		  
		  Thread.sleep(1000);
		  
		  Driver.driver.findElement(By.xpath("//div[@id='Clm_Presentation_vod__c_body']/table/tbody/tr[2]/th/a")).click();
		  
		  Thread.sleep(1000);
		  
		 // Driver.driver.findElement(By.xpath("//input[@autocomplete='off']")).sendKeys(presentationName);
		  
		  ////////
		
		 /* try{
			  
			  wUtils.waitForElementPresent("//a[span[img[@alt='CLM Presentation']]/following-sibling::strong[text()="+"'"+presentationName+"'"+"]]");
			  String actPresName=Driver.driver.findElement(By.xpath("//a[span[img[@alt='CLM Presentation']]/following-sibling::strong[text()="+"'"+presentationName+"'"+"]]")).getText();
			  
			  System.out.println("Presentation Name: "+actPresName.trim());
			  if(actPresName.trim().equals(presentationName)){
				  
				wUtils.waitForElementClickable("//a[span[img[@alt='CLM Presentation']]/following-sibling::strong[text()="+"'"+presentationName+"'"+"]]");
				Driver.driver.findElement(By.xpath("//a[span[img[@alt='CLM Presentation']]/following-sibling::strong[text()="+"'"+presentationName+"'"+"]]")).click();
				  
				 Reporter.log("Presentation found and seleceted", true);
				  
			  }
			  else{
				  throw new SkipException("Duplicate presentation found OR presentation doesn't exist");
			  }
			  try{}
		  }
		  catch(Exception e){
			  e.printStackTrace();
			  Reporter.log("Duplicate presentation found or presentation doesn't exist", true);
			  if (frmOpt == null) {
			        frmOpt = new JFrame();
			    }
			    frmOpt.setVisible(true);
			    frmOpt.setLocationRelativeTo(null);
			    frmOpt.setAlwaysOnTop(true);
				JOptionPane.showMessageDialog(frmOpt, "Duplicate presentation found or presentation doesn't exist");
			    Driver.driver.quit();
			    throw new SkipException("Duplicate presentation found or presentation doesn't exist"); 
		  }*/
		 
	      //wUtils.waitFortextPresent("//h2[contains(text(),"+"'"+presentationName+"'"+")]",presentationName);
	      
	      
	 	 
	     sandbox=Driver.driver.findElement(By.xpath("//span[contains(text(),'Sandbox')]/following-sibling::span")).getText(); 
          
	     if(sandbox.equals("NCLMRTEDF1")||sandbox.equals("NCLMRTEDF1")){
		  try{
		  wUtils.waitFortextPresent("//a[contains(text(),'Show')]", "Show");
		  if(Driver.driver.findElement(By.xpath("//a[contains(text(),'Show')]")).isDisplayed()){
	      Driver.driver.findElement(By.xpath("//a[contains(text(),'Show')]")).click();
	      Thread.sleep(1000);
	     
		  }
		  }
		  catch(Exception e){
			  
		  }
	     }
		  wUtils.captureScreenshot(new SimpleDateFormat("dd-MMM-yyyy_hh_mm_ss"), 0);
		  Reporter.log("Screenshot of presentation page is taken", true);
	  }
	else{
		Reporter.log("Multiple MVR files are present", true);
		JOptionPane.showMessageDialog(null, "Multiple MVR Input files are present");
		throw new SkipException("Multiple MVR files are present");
	}}
	    else{
	    	Reporter.log("MVR file mentioned in InputFile.xlsx is Missing", true);
	    	JOptionPane.showMessageDialog(null, "MVR file mentioned in InputFile.xlsx is Missing");
			throw new SkipException("MVR file mentioned in InputFile.xlsx is Missing");
	    	
	    }
}
	
  @Test(dependsOnMethods={"login"})
  public void clmPresentationMetadataReport() throws InvalidFormatException, IOException {
	 
	  
	  
	  slideCount=Driver.driver.findElements(By.xpath("//td[@class='actionColumn']/following-sibling::th/a"));
	 
	
	
	  //Against Content Map
	  
	  String clmPresName=Driver.driver.findElement(By.xpath("//td[contains(text(),'CLM Presentation Name')]/following-sibling::td[1]/div")).getText();
	  eLib.setexcelMasterMetaData(0, 3, 3, sLib.outputReportPath(), clmPresName);
	  
	  //Against MVR Report
	  String clmProduct=null;
	  try {
	   clmProduct=Driver.driver.findElement(By.xpath("//td[contains(text(),'Product')]/following-sibling::td[1]/div/a")).getText();
	  }
	  catch(Exception e)
	  {System.out.println("No product");}
	  
	  eLib.setexcelMasterMetaData(1, 3, 3, sLib.outputReportPath(), clmProduct);
	  
	  String clmPresentationID=Driver.driver.findElement(By.xpath("//td[span[contains(text(),'Presentation Id')]]/following-sibling::td[1]/div")).getText();
	  eLib.setexcelMasterMetaData(1, 4, 3, sLib.outputReportPath(), clmPresentationID);
	  
	  String clmExtID1=Driver.driver.findElement(By.xpath("//td[span[contains(text(),'External Id')]]/following-sibling::td[1]/div")).getText();
	  eLib.setexcelMasterMetaData(1, 5, 3, sLib.outputReportPath(), clmExtID1);
	  
	//Against Content Map
	   StringBuilder dOrder = new StringBuilder();
	   
	  //Against MVR Report
	   StringBuilder slideGuid = new StringBuilder();//to be removed
	   StringBuilder keyMsgProduct= new StringBuilder();
	   
	  for(int i=1,j=1;i<=slideCount.size();i++,j++){
		  
		//Against Content Map
		  try{
		  dOrder.append(Driver.driver.findElement(By.xpath("//tr[th[text()='Display Order']/preceding-sibling::th[2]]/following-sibling::tr/th/following-sibling::td[text()="+j+"]/preceding-sibling::th/a")).getText()+"\n");
		  }
		  catch(NoSuchElementException e){
				 
			  if (frmOpt == null) {
			        frmOpt = new JFrame();
			    }
			    frmOpt.setVisible(true);
			    frmOpt.setLocationRelativeTo(null);
			    frmOpt.setAlwaysOnTop(true);
				
				JOptionPane.showMessageDialog(frmOpt, "Display Order number "+j+" is missing in salesforce portal");
				File file=new File("C:\\MVR_Report_Automation\\Screenshots");
				File file1=new File("C:\\MVR_Report_Automation\\Output Report");
				FileUtils.cleanDirectory(file);
				FileUtils.cleanDirectory(file1);
				throw new SkipException("Display Order number "+j+" is missing in salesforce portal");
			}
		 
		  
        switch(sandbox)
        
        { 
        case "NCLMRTEDF1":
			  slideGuid.append(Driver.driver.findElement(By.xpath("//tr[th[text()='External ID']]/following-sibling::tr["+j+"]/td[text()="+j+"]/following-sibling::td[2]")).getText()+"\n");
			  keyMsgProduct.append(Driver.driver.findElement(By.xpath("//tr[th[text()='Key Message Product']]/following-sibling::tr["+j+"]/td[text()="+j+"]/following-sibling::td[1]")).getText()+"\n");
			  Reporter.log("Environment is NCLMRTEDF1", true);
			  break;
			  
        case "EUDF1":
			  slideGuid.append(Driver.driver.findElement(By.xpath("//tr[th[text()='External ID']]/following-sibling::tr["+j+"]/td[text()="+j+"]/following-sibling::td[2]")).getText()+"\n");
			  keyMsgProduct.append(Driver.driver.findElement(By.xpath("//td[contains(text(),'Product')]/following-sibling::td[1]/div/a")).getText()+"\n");
			  Reporter.log("Environment is EUDF1", true);
			  break;
			  
		  case "FctoryTest":
			  slideGuid.append(Driver.driver.findElement(By.xpath("//tr[th[text()='External ID']]/following-sibling::tr["+j+"]/td[text()="+j+"]/following-sibling::td[2]")).getText()+"\n");
			  keyMsgProduct.append(Driver.driver.findElement(By.xpath("//tr[th[text()='Key Message Product']]/following-sibling::tr["+j+"]/td[text()="+j+"]/following-sibling::td[1]")).getText()+"\n");
			  Reporter.log("Environment is CLMTestEnv", true);
			  break;
			  
		  case "CLMFactory":
			  slideGuid.append(Driver.driver.findElement(By.xpath("//tr[th[text()='External ID']]/following-sibling::tr["+j+"]/td[text()="+j+"]/following-sibling::td[1]")).getText()+"\n");
			  keyMsgProduct.append(Driver.driver.findElement(By.xpath("//td[contains(text(),'Product')]/following-sibling::td[1]/div/a")).getText()+"\n");
			  Reporter.log("Environment is CLMFactory", true);
			  break;
			  
		  
		  }

	  }
	
	  //Against Content Map
	  String clmDispOrder=dOrder.toString();
	  eLib.setexcelMasterMetaData(0, 4, 3,sLib.outputReportPath(), clmDispOrder);
	  
	  
	  String clmKeyMessage=dOrder.toString();
	  eLib.setexcelMasterMetaData(0, 5, 3,sLib.outputReportPath(), clmKeyMessage);
	  
	//Against MVR Report
	 String  clmExtID2=slideGuid.toString();//to be removed
	  String  clmkeyMsgProduct=keyMsgProduct.toString();
	eLib.setexcelMasterMetaData(1, 6, 3, sLib.outputReportPath(), clmExtID2);//to be removed
	  eLib.setexcelMasterMetaData(1, 7, 3, sLib.outputReportPath(), clmkeyMsgProduct);
	  
	  //General check
	  String hidden=Driver.driver.findElement(By.xpath("//td[contains(text(),'Hidden?')]/following-sibling::td[1]/div/img")).getAttribute("title");
	  if(hidden.equals("Not Checked")){
		  eLib.setexcelMasterMetaData(2, 3, 3, sLib.outputReportPath(), "Slide is not hidden");
		  eLib.setexcelMasterMetaDataColor(2, 3, 4, sLib.outputReportPath(), "PASS",HSSFColor.GREEN.index);
		   }
		  else{
			  eLib.setexcelMasterMetaData(2, 3, 3, sLib.outputReportPath(), "Slide is hidden");
			  eLib.setexcelMasterMetaDataColor(2, 3, 4, sLib.outputReportPath(), "FAIL",HSSFColor.RED.index);
			  
		  }
	  
	  String approved=Driver.driver.findElement(By.xpath("//td[span[contains(text(),'Approved?')]]/following-sibling::td[1]/div/img")).getAttribute("title");
	  if(approved.equals("Checked")){
		  eLib.setexcelMasterMetaData(2, 4, 3, sLib.outputReportPath(), "Slide is approved");
		  eLib.setexcelMasterMetaDataColor(2, 4, 4, sLib.outputReportPath(), "PASS",HSSFColor.GREEN.index);
		   }
		  else{
			  eLib.setexcelMasterMetaData(2, 4, 3, sLib.outputReportPath(), "Slide is not approved");
			  eLib.setexcelMasterMetaDataColor(2, 4, 4, sLib.outputReportPath(), "FAIL",HSSFColor.RED.index);
			  
		  }
	  
	  String iRepPres=Driver.driver.findElement(By.xpath("//td[span[contains(text(),'iREP Presentation')]]/following-sibling::td[1]/div/img")).getAttribute("title");
	  if(iRepPres.equals("Checked")){
		  eLib.setexcelMasterMetaData(2, 30, 3, sLib.outputReportPath(), "Checked");
		  eLib.setexcelMasterMetaDataColor(2, 30, 4, sLib.outputReportPath(), "PASS",HSSFColor.GREEN.index);
		   }
		  else{
			  eLib.setexcelMasterMetaData(2, 30, 3, sLib.outputReportPath(), "Not Checked");
			  eLib.setexcelMasterMetaDataColor(2, 30, 4, sLib.outputReportPath(), "FAIL",HSSFColor.RED.index);
			  
		  }
	  
	//Against Content Map

	    FileInputStream file = new FileInputStream(new File(sLib.outputReportPath()));
		HSSFWorkbook workbook = new HSSFWorkbook(file);
		HSSFSheet sheet = workbook.getSheetAt(0);
		HSSFCell cell = null;
		
		cell = sheet.getRow(3).getCell(4);
																							//1,  1
								                                //(String sheetName , int rowNum , int colNum,String excelpath
																			//(String SheetName ,int rowNum, int colNum,String excelpath)
		//String mvrPresExtName=eLib.getexcelMasterMetaData("Presentation", sLib.getCelldata("MVRP", 1, 0, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRP", 1, 2, SalesforceLib.INPUT_FILE),sLib.mvrReportPath());
		
    	
		String mvrPresExtName=eLib.getexcelMasterMetaData("Login", 2, 5, SalesforceLib.INPUT_FILE); 
		
		
	 	cell.setCellValue(mvrPresExtName);
	 	
	 	cell = sheet.getRow(4).getCell(4);
	 	
	 	totalRowCountMvrs=eLib.getLastRowNum(2, SalesforceLib.INPUT_FILE);
		System.out.println(totalRowCountMvrs);
	 	
		StringBuilder DispOrder = new StringBuilder();
		  for(int i=1;i<=totalRowCountMvrs;i++){
			  
			  DispOrder.append(eLib.getexcelMasterMetaData("Slides", sLib.getCelldata("MVRS", i, 1, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRS", i, 2, SalesforceLib.INPUT_FILE), sLib.mvrReportPath())+"\n");

		  }
		  String mvrSlidesIntName=DispOrder.toString();
		  cell.setCellValue(mvrSlidesIntName);
	
		
		cell = sheet.getRow(5).getCell(4);
		cell.setCellValue(mvrSlidesIntName);
		
		
		file.close();
	 	FileOutputStream outFile =new FileOutputStream(new File(sLib.outputReportPath()));
		workbook.write(outFile);
		
		ArrayList<String> al1=new ArrayList<String>();
	    ArrayList<String> al2=new ArrayList<String>();
	      
	      for(int i=3,j=3,k=4;i<=5;i++){
	      al1.add(eLib.getexcelMasterMetaData("Against Content map", i, j,sLib.outputReportPath()));//6,3
	 	  al2.add(eLib.getexcelMasterMetaData("Against Content map",i, k,sLib.outputReportPath()));//6,4
	 	 
	      }
	      
	      eLib.singleCellMultirowComparison(sLib.outputReportPath(), 0, 3, 5, al1, al2);
	      
	      
	 	 outFile.close();
	 	 workbook.close();
	 	
	 	//Against MVR Report

	 	    FileInputStream file1 = new FileInputStream(new File(sLib.outputReportPath()));
			HSSFWorkbook workbook1 = new HSSFWorkbook(file1);
			HSSFSheet sheet1 = workbook1.getSheetAt(1);
			HSSFCell cell1 = null;
			totalRowCountMvrs=eLib.getLastRowNum(2, SalesforceLib.INPUT_FILE);
			
		    String mvrPresentationGUID=eLib.getexcelMasterMetaData("Presentation", sLib.getCelldata("MVRP", 1, 0, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRP", 1, 6, SalesforceLib.INPUT_FILE), sLib.mvrReportPath());
			
		    String mvrPresProduct=null;
	        cell1 = sheet1.getRow(3).getCell(4, Row.CREATE_NULL_AS_BLANK);
	        try {
		  mvrPresProduct=eLib.getexcelMasterMetaData("Presentation", sLib.getCelldata("MVRP", 1, 0, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRP", 1, 4, SalesforceLib.INPUT_FILE), sLib.mvrReportPath());
	        } 
	        catch(Exception e)
	        {System.out.println("null value");}
		 	cell1.setCellValue(mvrPresProduct);
		 	
		 	cell1 = sheet1.getRow(4).getCell(4);
		 	cell1.setCellValue(mvrPresentationGUID);   
		 	
		 	
		 	cell1 = sheet1.getRow(5).getCell(4);
		 	String mvrPresPaidIntName=mvrPresentationGUID+"::"+eLib.getexcelMasterMetaData("Presentation", sLib.getCelldata("MVRP", 1, 0, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRP", 1, 1, SalesforceLib.INPUT_FILE), sLib.mvrReportPath());
		 	cell1.setCellValue(mvrPresPaidIntName);
				
			cell1 = sheet1.getRow(6).getCell(4);//to be removed
			StringBuilder prGuidSlideGuidVer = new StringBuilder();
		     for(int i=1;i<=totalRowCountMvrs;i++){
					  
			 prGuidSlideGuidVer.append(mvrPresentationGUID+"::"+eLib.getexcelMasterMetaData("Slides", sLib.getCelldata("MVRS", i, 1, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRS", i, 7, SalesforceLib.INPUT_FILE), sLib.mvrReportPath())+"::"+eLib.getexcelMasterMetaData("Slides", sLib.getCelldata("MVRS", i, 1, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRS", i, 6, SalesforceLib.INPUT_FILE), sLib.mvrReportPath())+"\n");
	        }
			String mvrPrGuidSlGuidSlVer=prGuidSlideGuidVer.toString();
			cell1.setCellValue(mvrPrGuidSlGuidSlVer);
				
				
				
		      cell1 = sheet1.getRow(7).getCell(4);
		      StringBuilder slidesProduct = new StringBuilder();
			  for(int i=1;i<=totalRowCountMvrs;i++){
				  try
				  {
				  slidesProduct.append(eLib.getexcelMasterMetaData("Slides", sLib.getCelldata("MVRS", i, 1, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRS", i, 5, SalesforceLib.INPUT_FILE), sLib.mvrReportPath())+"\n");
				  String mvrSlidesProduct=slidesProduct.toString();
			      cell1.setCellValue(mvrSlidesProduct);
				  }catch(Exception e)
			        {System.out.println("null value");}
			  }
			 
			
		
			 	
			file1.close();
			FileOutputStream outFile1 =new FileOutputStream(new File(sLib.outputReportPath()));
			workbook1.write(outFile1);
				
				ArrayList<String> al3=new ArrayList<String>();
			    ArrayList<String> al4=new ArrayList<String>();
			    try {
			      for(int i=3,j=3,k=4;i<=7;i++){
			      al3.add(eLib.getexcelMasterMetaData("Against MVR report", i, j,sLib.outputReportPath()));//6,3
			 	  al4.add(eLib.getexcelMasterMetaData("Against MVR report",i, k,sLib.outputReportPath()));//6,4
			 	 
			      }
			      
			      eLib.singleCellMultirowComparison(sLib.outputReportPath(), 1, 3, 5, al3, al4);
  					}
			      catch(Exception e)
			        {System.out.println("no comparision");}
			 	
				outFile1.close();
			 	workbook1.close();
	  
	  
	  
  }
  
 @Test (dependsOnMethods={"clmPresentationMetadataReport"})
  public void keyMessageMetadataReport() throws InvalidFormatException, IOException, InterruptedException {
	  
	 
	 slideCount=Driver.driver.findElements(By.xpath("//td[@class='actionColumn']/following-sibling::th/a"));
	 
	 
	 //Against MVR Report
	 String clmPresentationID=Driver.driver.findElement(By.xpath("//td[span[contains(text(),'Presentation Id')]]/following-sibling::td[1]/div")).getText();
	 
	//Against Content Map
	 
	  StringBuilder message = new StringBuilder();
	  StringBuilder desc = new StringBuilder();
	  StringBuilder PresName = new StringBuilder();
	  
	//Against MVR Report
	  StringBuilder slideVer = new StringBuilder();
	  StringBuilder slideDesc = new StringBuilder();
	  StringBuilder clmID = new StringBuilder();
	  StringBuilder webview = new StringBuilder();
	  
	  StringBuilder extID1 = new StringBuilder();
	  StringBuilder product = new StringBuilder();
	  StringBuilder extID2 = new StringBuilder();
	  //General check
	  StringBuilder mediaFileName = new StringBuilder();
	  StringBuilder PresSlideName = new StringBuilder();
	  StringBuilder active = new StringBuilder();
	  StringBuilder genCheckmessage = new StringBuilder();//=========clmPresentation2====//
	  StringBuilder mediaFileCrc = new StringBuilder();
	  StringBuilder lastModifiedBy = new StringBuilder();
	  StringBuilder MsgNameLineBreak = new StringBuilder();
	  StringBuilder MsgDescLineBreak = new StringBuilder();
	  for(int i=1,j=1;i<=slideCount.size();i++,j++){
		  
		  Driver.driver.findElement(By.xpath("//tr[th[text()='Display Order']/preceding-sibling::th[2]]/following-sibling::tr/th/following-sibling::td[text()="+j+"]/preceding-sibling::th/a")).click();
		   
		 
		  Thread.sleep(1000);
		  //==============================================
		  wUtils.captureScreenshot(new SimpleDateFormat("dd-MMM-yyyy_hh_mm_ss"), i);
		//Against Content Map
		  message.append(Driver.driver.findElement(By.xpath("//td[contains(text(),'Message')]/following-sibling::td[1]/div")).getText()+"\n");
		  desc.append(Driver.driver.findElement(By.xpath("//tbody//tr[3]//following-sibling::td[1]//div[@id]")).getText()+"\n");
		  PresName.append(Driver.driver.findElement(By.xpath("//tr[th[text()='CLM Presentation']]/following-sibling::tr/td[2]/a")).getText()+"\n");
		//Against MVR Report
		  slideVer.append(Driver.driver.findElement(By.xpath("//td[contains(text(),'Slide Version')]/following-sibling::td[1]/div")).getText()+"\n");
		  slideDesc.append(Driver.driver.findElement(By.xpath("//td[span[contains(text(),'Slide Description')]]/following-sibling::td[1]/div")).getText()+"\n");
		  clmID.append(Driver.driver.findElement(By.xpath("//td[span[contains(text(),'CLM ID')]]/following-sibling::td[1]/div")).getText()+"\n");
		 webview.append(Driver.driver.findElement(By.xpath("//tr//td[contains(text(),'iOS Viewer')]//following-sibling::td[1]//div[@id]")).getText()+"\n");
		  
		  extID1.append(Driver.driver.findElement(By.xpath("//td[span[contains(text(),'External Id')]]/following-sibling::td[1]/div")).getText()+"\n");
		
		  try {
		  product.append(Driver.driver.findElement(By.xpath("//td[contains(text(),'Product')]/following-sibling::td[1]/div/a")).getText()+"\n");
		 
		  }catch(Exception e)
		  {System.out.println("No product");}
		  
		  extID2.append(clmPresentationID+"::"+Driver.driver.findElement(By.xpath("//td[span[contains(text(),'External Id')]]/following-sibling::td[1]/div")).getText()+"\n");
		  //General Check
		  mediaFileName.append(Driver.driver.findElement(By.xpath("//td[contains(text(),'Media File Name')]/following-sibling::td[1]/div")).getText()+"\n");
		  PresSlideName.append(Driver.driver.findElement(By.xpath("//tr[th[text()='CLM Presentation']]/following-sibling::tr/th/a")).getText()+"\n");
		  String status= Driver.driver.findElement(By.xpath("//td[contains(text(),'Active')]/following-sibling::td[1]/div/img")).getAttribute("title");
		  if(status.equals("Checked")){
			  active.append("Slide is Active"+"\n");
			  eLib.setexcelMasterMetaDataColor(2, 11, 4, sLib.outputReportPath(), "PASS",HSSFColor.GREEN.index);
			   }
			  else{
				  active.append("Slide is Inactive"+"\n");
				  eLib.setexcelMasterMetaDataColor(2, 11, 4, sLib.outputReportPath(), "FAIL",HSSFColor.RED.index);
				  
			  }
		  
		  String nameLineBreak= Driver.driver.findElement(By.xpath("//td[contains(text(),'Message')]/following-sibling::td[1]/div")).getText();
		  if(nameLineBreak.contains("\n")){
			  MsgNameLineBreak.append("There are line breaks"+"\n");
			  eLib.setexcelMasterMetaDataColor(2, 28, 4, sLib.outputReportPath(), "FAIL",HSSFColor.RED.index);
			   }
			  else{
				  MsgNameLineBreak.append("There are no line breaks"+"\n");
				  eLib.setexcelMasterMetaDataColor(2, 28, 4, sLib.outputReportPath(), "PASS",HSSFColor.GREEN.index);
				  
			  }
		  
		  String descLineBreak= Driver.driver.findElement(By.xpath("//td[span[text()='Description']]/following-sibling::td[1]/div")).getText();
		  if(descLineBreak.contains("\n")){
			  MsgDescLineBreak.append("There are line breaks"+"\n");
			  eLib.setexcelMasterMetaDataColor(2, 29, 4, sLib.outputReportPath(), "FAIL",HSSFColor.RED.index);
			   }
			  else{
				  MsgDescLineBreak.append("There are no line breaks"+"\n");
				  eLib.setexcelMasterMetaDataColor(2, 29, 4, sLib.outputReportPath(), "PASS",HSSFColor.GREEN.index);
				  
			  }
		  genCheckmessage.append(Driver.driver.findElement(By.xpath("//td[contains(text(),'Message')]/following-sibling::td[1]/div")).getText()+"\n");//=========clmPresentation2/3====//
		  mediaFileCrc.append(Driver.driver.findElement(By.xpath("//td[contains(text(),'Media File CRC')]/following-sibling::td[1]/div")).getText()+"\n");
		  lastModifiedBy.append(Driver.driver.findElement(By.xpath("//td[contains(text(),'Last Modified By')]/following-sibling::td[1]/div")).getText()+"\n");
		  Driver.driver.navigate().back();
		  wUtils.implicitWait();
		  
		  
		//Show More button
		  if(sandbox.equals("NCLMRTEDF1")||sandbox.equals("CLMDev")){
		  try{
		  wUtils.waitFortextPresent("//a[contains(text(),'Show')]", "Show");
		  if(Driver.driver.findElement(By.xpath("//a[contains(text(),'Show')]")).isDisplayed()){
	      Driver.driver.findElement(By.xpath("//a[contains(text(),'Show')]")).click();
	      Thread.sleep(1000);
	      	     
		  }
		  }
		  catch(Exception e){
			  
		  }
		  }
		  
		 
	  }
	//Against Content Map
	  String keyMessage=message.toString();
	  String keyDescription=desc.toString();
	  String keyDispOrder=message.toString();
	  String KeyPresName=PresName.toString();
	//Against MVR Report
	  String keySlidever=slideVer.toString();
	  String keyslideDesc=slideDesc.toString();  
	  String keyclmID=clmID.toString();
	  String keywebview=webview.toString();
	  
	  
	  String keyExtId1=extID1.toString();
	  String keyProduct=product.toString();
	  String  keyExtID2=extID2.toString();
	//General check
	  String keyMediaFileName=mediaFileName.toString();
	  String slideActive=active.toString();
	  String KeyPresSlideName=PresSlideName.toString();
	  String clmMessage=genCheckmessage.toString();//=========clmPresentation2====//
	  String clmMediaFileCrc=mediaFileCrc.toString();
	  String clmMediaFileName=mediaFileName.toString();
	  String clmLastModifiedBy=lastModifiedBy.toString();
	  String keyMsgNameLineBreak=MsgNameLineBreak.toString();
	  String keyMsgDescLineBreak=MsgDescLineBreak.toString();
	//Against Content Map
	  eLib.setexcelMasterMetaData(0, 10, 3, sLib.outputReportPath(), keyMessage);
	  eLib.setexcelMasterMetaData(0, 11, 3, sLib.outputReportPath(), keyDescription);
	  eLib.setexcelMasterMetaData(0, 12, 3, sLib.outputReportPath(), keyDispOrder);
	  eLib.setexcelMasterMetaData(0, 13, 3, sLib.outputReportPath(), KeyPresName);
	//Against MVR Report
	  
	  eLib.setexcelMasterMetaData(1, 12, 3, sLib.outputReportPath(), keySlidever);
	  eLib.setexcelMasterMetaData(1, 13, 3, sLib.outputReportPath(), keyslideDesc);
	  eLib.setexcelMasterMetaData(1, 14, 3, sLib.outputReportPath(), keyclmID);
	  
	  
	  eLib.setexcelMasterMetaData(1, 15, 3, sLib.outputReportPath(), keyExtId1);
	  eLib.setexcelMasterMetaData(1, 16, 3, sLib.outputReportPath(), keyProduct);
	  eLib.setexcelMasterMetaData(1, 17, 3, sLib.outputReportPath(), keyExtID2);
	  eLib.setexcelMasterMetaData(1, 18, 3, sLib.outputReportPath(), keyProduct);
	  
	//General Check
	  eLib.setexcelMasterMetaData(2, 10, 3, sLib.outputReportPath(), keyMediaFileName);
	  eLib.setexcelMasterMetaData(2, 11, 3, sLib.outputReportPath(), slideActive);
	  eLib.setexcelMasterMetaData(2, 12, 3, sLib.outputReportPath(), KeyPresSlideName);
	  eLib.setexcelMasterMetaData(2, 13, 3, sLib.outputReportPath(), keywebview); // change number
	  
	  eLib.setexcelMasterMetaData(2, 34, 1, sLib.outputReportPath(), clmMessage);//=========clmPresentation2====//
	  eLib.setexcelMasterMetaData(2, 34, 2, sLib.outputReportPath(), clmMediaFileCrc);
	  eLib.setexcelMasterMetaData(2, 34, 3, sLib.outputReportPath(), clmMediaFileName);
	  eLib.setexcelMasterMetaData(2, 34, 4, sLib.outputReportPath(), clmLastModifiedBy);
	  eLib.setexcelMasterMetaData(2, 28, 3, sLib.outputReportPath(), keyMsgNameLineBreak);
	  eLib.setexcelMasterMetaData(2, 29, 3, sLib.outputReportPath(), keyMsgDescLineBreak); 
	  
	  eLib.setexcelMasterMetaData(2, 5, 3, sLib.outputReportPath(), KeyPresSlideName);//=========clmPresentation1====//
	  ArrayList<String> PresSlideNameResult=new ArrayList<String>();
	  PresSlideNameResult.add(eLib.getexcelMasterMetaData("General check", 5, 3,sLib.outputReportPath()));
	  eLib.allDifferentPass(sLib.outputReportPath(), 2, 5, 4, PresSlideNameResult);
	  eLib.allDifferentPass(sLib.outputReportPath(), 2, 12, 4, PresSlideNameResult);
	  ArrayList<String> MediaFileNameResult=new ArrayList<String>();
	  MediaFileNameResult.add(eLib.getexcelMasterMetaData("General check", 10, 3,sLib.outputReportPath()));
	  eLib.allDifferentPass(sLib.outputReportPath(), 2, 10, 4, MediaFileNameResult);
	  eLib.setexcelMasterMetaData(2, 34, 5, sLib.outputReportPath(), "PASS");
	  

	  
	//Against Content Map
	  
	    FileInputStream file = new FileInputStream(new File(sLib.outputReportPath()));
		HSSFWorkbook workbook = new HSSFWorkbook(file);
		HSSFSheet sheet = workbook.getSheetAt(0);
		HSSFCell cell = null;
		totalRowCountMvrs=eLib.getLastRowNum(2, SalesforceLib.INPUT_FILE);
		
		cell = sheet.getRow(10).getCell(4);
		StringBuilder DispOrder = new StringBuilder();
		  for(int i=1;i<=totalRowCountMvrs;i++){
			  
			  DispOrder.append(eLib.getexcelMasterMetaData("Slides", sLib.getCelldata("MVRS", i, 1, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRS", i, 2, SalesforceLib.INPUT_FILE), sLib.mvrReportPath())+"\n");

		  }
		  String mvrSlidesIntName=DispOrder.toString();
		  cell.setCellValue(mvrSlidesIntName);
		
		cell = sheet.getRow(11).getCell(4);
		StringBuilder description = new StringBuilder();
		  for(int i=1;i<=totalRowCountMvrs;i++){
			  
			  description.append(eLib.getexcelMasterMetaData("Slides", sLib.getCelldata("MVRS", i, 1, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRS", i, 3, SalesforceLib.INPUT_FILE), sLib.mvrReportPath())+"\n");

		  }
		  String mvrdescription=description.toString();
		  cell.setCellValue(mvrdescription);
		
		cell = sheet.getRow(12).getCell(4);
		cell.setCellValue(mvrSlidesIntName);
		
		cell = sheet.getRow(13).getCell(4);
		StringBuilder presExtName = new StringBuilder();
		for(int i=1;i<=totalRowCountMvrs;i++){
		presExtName.append(eLib.getexcelMasterMetaData("Presentation", sLib.getCelldata("MVRP", 1, 0, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRP", 1, 2, SalesforceLib.INPUT_FILE), sLib.mvrReportPath())+"\n");
		
		}
		String mvrPresExtName=presExtName.toString();
		cell.setCellValue(mvrPresExtName);
		
		file.close();
	 	FileOutputStream outFile =new FileOutputStream(new File(sLib.outputReportPath()));
		workbook.write(outFile);
		
		
		ArrayList<String> al1=new ArrayList<String>();
	    ArrayList<String> al2=new ArrayList<String>();
	      
	      for(int i=10,j=3,k=4;i<=13;i++){
	      al1.add(eLib.getexcelMasterMetaData("Against Content map", i, j,sLib.outputReportPath()));//6,3
	 	  al2.add(eLib.getexcelMasterMetaData("Against Content map",i, k,sLib.outputReportPath()));//6,4
	 	 
	      }
	      
	      eLib.singleCellMultirowComparison(sLib.outputReportPath(), 0, 10, 5, al1, al2);
	      
	 	
        outFile.close();
	 	workbook.close();
	 	
		//Against MVR Report
	 	
	 	FileInputStream file2 = new FileInputStream(new File(sLib.outputReportPath()));
		HSSFWorkbook workbook2 = new HSSFWorkbook(file2);
		HSSFSheet sheet2 = workbook2.getSheetAt(1);
		HSSFCell cell2 = null;
		HSSFCell cell3 = null;
		totalRowCountMvrs=eLib.getLastRowNum(2, SalesforceLib.INPUT_FILE);
		String mvrPresentationGUID=eLib.getexcelMasterMetaData("Presentation", sLib.getCelldata("MVRP", 1, 0, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRP", 1, 6, SalesforceLib.INPUT_FILE), sLib.mvrReportPath());

		
	 	cell2 = sheet2.getRow(12).getCell(4);
	 	StringBuilder slideVersion = new StringBuilder();
		  for(int i=1;i<=totalRowCountMvrs;i++){
			  
			  slideVersion.append(eLib.getexcelMasterMetaData("Slides", sLib.getCelldata("MVRS", i, 1, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRS", i, 6, SalesforceLib.INPUT_FILE), sLib.mvrReportPath())+"\n");

		  }
		  String keySlideVer=slideVersion.toString();
	 	  cell2.setCellValue(keySlideVer);
	 	
	 	cell2 = sheet2.getRow(13).getCell(4);
	 	StringBuilder slideIntName = new StringBuilder();
		  for(int i=1;i<=totalRowCountMvrs;i++){
			  
			  slideIntName.append(eLib.getexcelMasterMetaData("Slides", sLib.getCelldata("MVRS", i, 1, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRS", i, 2, SalesforceLib.INPUT_FILE), sLib.mvrReportPath())+"\n");

		  }
		  String keySlideIntName=slideIntName.toString();
	 	  cell2.setCellValue(keySlideIntName);
	 	
	 	cell2 = sheet2.getRow(14).getCell(4);
	 	StringBuilder slideGuid = new StringBuilder();
		  for(int i=1;i<=totalRowCountMvrs;i++){
			  
			  slideGuid.append(eLib.getexcelMasterMetaData("Slides", sLib.getCelldata("MVRS", i, 1, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRS", i, 7, SalesforceLib.INPUT_FILE), sLib.mvrReportPath())+"\n");

		  }
		  String mvrSlideGuid=slideGuid.toString();
		  cell2.setCellValue(mvrSlideGuid);
			
		cell2 = sheet2.getRow(15).getCell(4);
		StringBuilder slideGUIDSlideVer = new StringBuilder();
		  for(int i=1;i<=totalRowCountMvrs;i++){
			  
			  slideGUIDSlideVer.append(eLib.getexcelMasterMetaData("Slides", sLib.getCelldata("MVRS", i, 1, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRS", i, 7, SalesforceLib.INPUT_FILE), sLib.mvrReportPath())+"::"+eLib.getexcelMasterMetaData("Slides", sLib.getCelldata("MVRS", i, 1, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRS", i, 6, SalesforceLib.INPUT_FILE), sLib.mvrReportPath())+"\n");

		  }
		  String mvrSlideGUIDSlideVer=slideGUIDSlideVer.toString();
	      cell2.setCellValue(mvrSlideGUIDSlideVer);
	    
	    
	    try {
	    StringBuilder slidesProduct = new StringBuilder();
		  for(int i=1;i<=totalRowCountMvrs;i++){
			  
			  slidesProduct.append(eLib.getexcelMasterMetaData("Slides", sLib.getCelldata("MVRS", i, 1, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRS", i, 5, SalesforceLib.INPUT_FILE), sLib.mvrReportPath())+"\n");

			  
			  
		  String mvrSlidesProduct=slidesProduct.toString();
		  cell2 = sheet2.getRow(16).getCell(4);  
		  cell2.setCellValue(mvrSlidesProduct);
		  cell2 = sheet2.getRow(18).getCell(4);
			cell2.setCellValue(mvrSlidesProduct);
		  }
	    }
	    catch(Exception e)
	    {}
	    
	    cell2 = sheet2.getRow(17).getCell(4);
	    StringBuilder prGuidSlideGuidVer = new StringBuilder();//here to be added
		  for(int i=1;i<=totalRowCountMvrs;i++){
			  
			  prGuidSlideGuidVer.append(mvrPresentationGUID+"::"+eLib.getexcelMasterMetaData("Slides", sLib.getCelldata("MVRS", i, 1, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRS", i, 7, SalesforceLib.INPUT_FILE), sLib.mvrReportPath())+"::"+eLib.getexcelMasterMetaData("Slides", sLib.getCelldata("MVRS", i, 1, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRS", i, 6, SalesforceLib.INPUT_FILE), sLib.mvrReportPath())+"\n");

		  }
		  String mvrPrGuidSlGuidSlVer=prGuidSlideGuidVer.toString();
		  cell2.setCellValue(mvrPrGuidSlGuidSlVer);
	    
			
		/*cell2 = sheet2.getRow(18).getCell(4);
		cell2.setCellValue(mvrSlidesProduct);*/	
				
			
		 	
		file2.close();
		FileOutputStream outFile2 =new FileOutputStream(new File(sLib.outputReportPath()));
		workbook2.write(outFile2);

			
		  ArrayList<String> al3=new ArrayList<String>();
	      ArrayList<String> al4=new ArrayList<String>();
	      
	      for(int i=12,j=3,k=4;i<=18;i++){
	      al3.add(eLib.getexcelMasterMetaData("Against MVR report", i, j,sLib.outputReportPath()));//6,3
	 	  al4.add(eLib.getexcelMasterMetaData("Against MVR report",i, k,sLib.outputReportPath()));//6,4
	 	 
	      }
	     
	      eLib.singleCellMultirowComparison(sLib.outputReportPath(), 1, 12, 5, al3, al4);
	      
	      
	 	
		outFile2.close();
	 	workbook2.close();

		
		
  
  }
 
 //@Test
 public void clickStreamReport() throws InvalidFormatException, IOException {
	  
 }
 
@Test(dependsOnMethods={"keyMessageMetadataReport"})
	public void sendEmailTest() throws InvalidFormatException, IOException {
	  
	  String from_User=eLib.getexcelMasterMetaData("Email", 1, 0, SalesforceLib.INPUT_FILE);  
	  String to_User=eLib.getexcelMasterMetaData("Email", 1, 1, SalesforceLib.INPUT_FILE);
	  String cc_User=eLib.getexcelMasterMetaData("Email", 1, 2, SalesforceLib.INPUT_FILE);
	  String presentationName=eLib.getexcelMasterMetaData("Login", 2, 5, SalesforceLib.INPUT_FILE);
	  String productName="unbrannded";
	  try {
	   productName=eLib.getexcelMasterMetaData("Presentation", sLib.getCelldata("MVRP", 1, 0, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRP", 1, 4, SalesforceLib.INPUT_FILE), sLib.mvrReportPath());
	  
	  }
	  catch(Exception e)
	  {}
	  String presentationID=eLib.getexcelMasterMetaData("Presentation", sLib.getCelldata("MVRP", 1, 0, SalesforceLib.INPUT_FILE), sLib.getCelldata("MVRP", 1, 6, SalesforceLib.INPUT_FILE), sLib.mvrReportPath());
	  
	  
	  if(actStatus.status1().contains("FAIL")||actStatus.status2().contains("FAIL")||actStatus.status3().contains("FAIL")){
			status="<html><font color='red'>FAIL</font></html>";
		    subjectStatus="FAIL";
		}
		else{
			status="<html><font color='green'>PASS</font></html>";
			subjectStatus="PASS";
		}
		System.out.println(subjectStatus);
	  
	  String htmlMessage="Team,<br><br> Validation team has successfully completed all planned operational & analytical test scripts for the <b>"+presentationName+"</b> iDetail as per the input MVR report.</br><br>Below is the summary:</br><br>";

			htmlMessage+="<table style=width:50% border=1 cellspacing=0 cellpadding=0><tr><td><b>iDetail Name</b></td> <td>"+presentationName+"</td></tr><tr><td><b>Product Name</b></td> <td>"+productName+"</td></tr><tr><td><b>Presentation ID</b></td> <td>"+presentationID+"</td></tr></tbody></table>";
            htmlMessage+="<br><table style=width:50% border=1 cellspacing=0 cellpadding=0><tbody ><tr><td><b>Environment</b></td> <td colspan=2>"+sandbox+"</td></tr><tr><td><b>Testing Parameters</b></td> <td><b>Operational Testing</b></td><td><b>Analytical Testing</b></td></tr><tr><td><b>STATUS</b></td> <td>"+status+"</td><td>"+status+"</td></tr></tbody></table>";
            htmlMessage+="</br><br>Find the email attachment for the detailed MVR Test Output report.</br><br><b>Note:</b> This is an automated mail. Do not reply to this mail.</br><br>Regards,<br>Validation Team</br>";
	  SendMail_PostmarkApp mail=new SendMail_PostmarkApp(from_User,to_User,"MVR Execution Report: "+presentationName+" "+"["+subjectStatus+"]",htmlMessage ,cc_User,false);
		 	  
	  mail.sendAttachmentMail();
		
	}
 
 @AfterTest
 public void logout(){
	  Driver.driver.quit();
	  //recorder.stop();
 }
 
 
}
