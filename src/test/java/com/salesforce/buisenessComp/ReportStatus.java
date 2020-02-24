package com.salesforce.buisenessComp;

import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;



import com.salesforce.buisenessComp.SalesforceLib;
import com.salesforce.genericLib.ExcelLib;

public class ReportStatus {
	
	ExcelLib eLib=new ExcelLib();
	SalesforceLib sLib=new SalesforceLib();
	
   public String status1() throws InvalidFormatException, IOException {
		
		StringBuilder agContentMapStatus = new StringBuilder();
		  for(int i=3;i<=5;i++){
			  
			  agContentMapStatus.append(eLib.getexcelMasterMetaData("Against Content map", i,5, sLib.outputReportPath())+"\n");

		  }
          for(int i=10;i<=13;i++){
			  
			  agContentMapStatus.append(eLib.getexcelMasterMetaData("Against Content map", i,5, sLib.outputReportPath())+"\n");

		  }
          
          agContentMapStatus.append(eLib.getexcelMasterMetaData("Against Content map", 18,5, sLib.outputReportPath())+"\n");
          
		  String status1=agContentMapStatus.toString();
		  return status1;
		

	}
	
    public String status2() throws InvalidFormatException, IOException {
		
		StringBuilder agMvrReportStatus = new StringBuilder();
		  for(int i=3;i<=7;i++){
			  
			  agMvrReportStatus.append(eLib.getexcelMasterMetaData("Against MVR report", i,5, sLib.outputReportPath())+"\n");

		  }
          for(int i=12;i<=18;i++){
			  
        	  agMvrReportStatus.append(eLib.getexcelMasterMetaData("Against MVR report", i,5, sLib.outputReportPath())+"\n");

		  }
          
         for(int i=23;i<=25;i++){
			  
        	  agMvrReportStatus.append(eLib.getexcelMasterMetaData("Against MVR report", i,5, sLib.outputReportPath())+"\n");

		  }
		  String status2=agMvrReportStatus.toString();
		  return status2;
		
		

	}
   
    public String status3() throws InvalidFormatException, IOException {
	
	StringBuilder genCheckStatus = new StringBuilder();
	  for(int i=3;i<=5;i++){
		  
		  genCheckStatus.append(eLib.getexcelMasterMetaData("General check", i,4, sLib.outputReportPath())+"\n");

	  }
      for(int i=10;i<=12;i++){
		  
    	  genCheckStatus.append(eLib.getexcelMasterMetaData("General check", i,4, sLib.outputReportPath())+"\n");

	  }
     for(int i=28;i<=30;i++){
		  
    	  genCheckStatus.append(eLib.getexcelMasterMetaData("General check", i,4, sLib.outputReportPath())+"\n");

	  }
     
     for(int i=17;i<=23;i++){
		  
   	  genCheckStatus.append(eLib.getexcelMasterMetaData("General check", i,4, sLib.outputReportPath())+"\n");

	  }

	  genCheckStatus.append(eLib.getexcelMasterMetaData("General check", 34,5, sLib.outputReportPath())+"\n");


	  String status3=genCheckStatus.toString();
	
	  return status3;
	}
	

}
