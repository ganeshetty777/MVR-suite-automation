package com.salesforce.genericLib;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.salesforce.buisenessComp.SalesforceLib;

public class SendMail_PostmarkApp {
	
	ExcelLib eLib=new ExcelLib();
	SalesforceLib sLib=new SalesforceLib();
	
	private String from;
	private String to;
	private String sub;
	private String body;
	private String cc;
	private boolean flag;
	
	public SendMail_PostmarkApp(String from,String to,String subject,String body,String cc,boolean fl) {
		this.body=body;
		this.from=from;
		this.sub=subject;
		this.to=to;
		this.cc=cc;
		this.flag=fl;
	}
	
	public SendMail_PostmarkApp(String from,String to,String subject,String body,boolean fl) {
		this.body=body;
		this.from=from;
		this.sub=subject;
		this.to=to;
		this.flag=fl;
	}
	
	
	
	public String sendAttachmentMail() throws InvalidFormatException, IOException{   
		 
		 
			String res="";
			
			Properties mailServerProperties = new Properties();
	        mailServerProperties.put("mail.smtp.host","smtp.postmarkapp.com");
	        mailServerProperties.put ("mail.smtp.port","587");
	        mailServerProperties.put ("mail.smtp.from",this.from);
	        mailServerProperties.setProperty("mail.smtp.auth","true");
			mailServerProperties.put("mail.smtp.starttls.enable", "true");
			mailServerProperties.put("mail.smtp.starttls.required", "true");
	        mailServerProperties.put("java.net.preferIPv4Stack" , "true");
	        System.setProperty("java.net.preferIPv4Stack" , "true");

	        Session session = Session.getInstance ( mailServerProperties );
	        
	        
	        
	        
	        MimeMessage messageToSend = new MimeMessage(session);
	        BodyPart messageBodyPart1 = new MimeBodyPart();
	        MimeBodyPart messageBodyPart2 = new MimeBodyPart();


	        try
	        {
	        	if(this.flag==false){
	        		String[] ccRecipientList = this.cc.split(";");
	        		String[] toRecipientList = this.to.split(";");
	            	InternetAddress[] ccRecipientAddress = new InternetAddress[ccRecipientList.length];
	            	InternetAddress[] toRecipientAddress = new InternetAddress[toRecipientList.length];
	            	int ccCounter = 0;
	            	int toCounter = 0;
	            	List<InternetAddress> list = new ArrayList<InternetAddress>();
	            	//list.add(new InternetAddress(this.to));
	            	for (String ccRecipient : ccRecipientList) {
	            		ccRecipientAddress[ccCounter] = new InternetAddress(ccRecipient.trim());
	//            	    System.out.println( ccRecipientAddress[ccCounter]);
	            	    list.add(ccRecipientAddress[ccCounter]);
	            	    ccCounter++;
	            	}
	            	for (String toRecipient : toRecipientList) {
	            		toRecipientAddress[toCounter] = new InternetAddress(toRecipient.trim());
	 //           	    System.out.println( toRecipientAddress[toCounter]);
	            	    list.add(toRecipientAddress[toCounter]);
	            	    toCounter++;
	            	}
	                messageToSend.setFrom ( new InternetAddress(this.from));
	                messageToSend.addRecipients( Message.RecipientType.TO,toRecipientAddress );
	                messageToSend.addRecipients( Message.RecipientType.CC,ccRecipientAddress );
//	                messageToSend.addRecipient(Message.RecipientType.BCC,new InternetAddress(this.from));
	               // InternetAddress[] add = {new InternetAddress(this.to)};
	                InternetAddress [] add = list.toArray(new InternetAddress[list.size()]);
	                messageToSend.setSubject(this.sub);
	                
	                //messageBodyPart1.setText(this.body); 
	                messageBodyPart1.setContent(this.body,"text/html");
	                String filename = sLib.outputReportPath();
	                DataSource source = new FileDataSource(filename);  
	    		    messageBodyPart2.setDataHandler(new DataHandler(source));  
	    		    messageBodyPart2.setFileName(filename);
	    		    
	    		    Multipart multipart = new MimeMultipart();  
	    		    multipart.addBodyPart(messageBodyPart1);  
	    		    multipart.addBodyPart(messageBodyPart2); 
	    		    
	    		    messageToSend.setContent(multipart );
	                
	  
	                Transport tr = session.getTransport("smtp");
	                tr.connect("0beb7054-395b-4085-984f-671d8fb4adb5","0beb7054-395b-4085-984f-671d8fb4adb5");
		             messageToSend.saveChanges(); 
		             tr.sendMessage(messageToSend,add);
	        	}
	        	else {
	        		String[] toRecipientList = this.to.split(";");
	            	InternetAddress[] toRecipientAddress = new InternetAddress[toRecipientList.length];
	            	int toCounter = 0;
	            	List<InternetAddress> list = new ArrayList<InternetAddress>();
	            	for (String toRecipient : toRecipientList) {
	            		toRecipientAddress[toCounter] = new InternetAddress(toRecipient.trim());
	//            	    System.out.println( toRecipientAddress[toCounter]);
	            	    list.add(toRecipientAddress[toCounter]);
	            	    toCounter++;
	            	}
	                messageToSend.setFrom ( new InternetAddress(this.from));
	                messageToSend.addRecipients( Message.RecipientType.TO,toRecipientAddress );
	                InternetAddress [] add = list.toArray(new InternetAddress[list.size()]);
	                messageToSend.setSubject(this.sub);
	                
	                messageBodyPart1.setText(this.body); 
	                String filename = sLib.outputReportPath();
	                DataSource source = new FileDataSource(filename);  
	    		    messageBodyPart2.setDataHandler(new DataHandler(source));  
	    		    messageBodyPart2.setFileName(filename);
	    		    
	    		    Multipart multipart = new MimeMultipart();  
	    		    multipart.addBodyPart(messageBodyPart1);  
	    		    multipart.addBodyPart(messageBodyPart2); 
	    		    
	    		    messageToSend.setContent(multipart);
	                
	                
	                Transport tr = session.getTransport("smtp");
	                tr.connect("0beb7054-395b-4085-984f-671d8fb4adb5","0beb7054-395b-4085-984f-671d8fb4adb5");
			        messageToSend.saveChanges(); 
			        tr.sendMessage(messageToSend,add);
	        	}
	        	      

	          res="Sent";
	         System.out.println(res);
	        }
	        catch (MessagingException ex)
	        {
	            ex.printStackTrace();
	            res="Cannot Send Mail.";
	            System.out.println(res);
	        }
			
			return res;
		
		   
		  
	}

}
