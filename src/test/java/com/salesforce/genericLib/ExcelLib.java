package com.salesforce.genericLib;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import src.com.salesforce.genericLib.WebDriverCommonUtils;



public class ExcelLib {
	
	
	
   public String getexcelMasterMetaData(String sheetName , int rowNum , int colNum,String excelpath) throws InvalidFormatException, IOException{
		   String data;
		   FileInputStream fis = new FileInputStream(excelpath);
			Workbook wb = WorkbookFactory.create(fis);		
			Sheet sh  = wb.getSheet(sheetName);
			Row row = sh.getRow(rowNum);
			Cell cell=row.getCell(colNum);
			
	        if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
	        	int i = (int)cell.getNumericCellValue();
	        	data = String.valueOf(i);
	        	} else {
	        	data = cell.toString();
	        	}
	        return data;
		}
	
	public void setexcelMasterMetaData(int sheetNo , int rowNum , int colNum,String excelpath,String data) throws InvalidFormatException, IOException{
		FileInputStream file = new FileInputStream(new File(excelpath));
		HSSFWorkbook workbook = new HSSFWorkbook(file);
		HSSFSheet sheet = workbook.getSheetAt(sheetNo);
		HSSFCell cell = null;
        cell = sheet.getRow(rowNum).getCell(colNum);
	 	cell.setCellValue(data);
	 	FileOutputStream outFile =new FileOutputStream(new File(excelpath));
		workbook.write(outFile);
		outFile.close();
	 	workbook.close();
		}
	
	public void setexcelMasterMetaDataColor(int sheetNo , int rowNum , int colNum,String excelpath,String data,short color) throws InvalidFormatException, IOException{
		FileInputStream file = new FileInputStream(new File(excelpath));
		HSSFWorkbook workbook = new HSSFWorkbook(file);
		HSSFSheet sheet = workbook.getSheetAt(sheetNo);
		HSSFCell cell = null;
        cell = sheet.getRow(rowNum).getCell(colNum);
	 	cell.setCellValue(data);
	 	HSSFCellStyle style = workbook.createCellStyle();
		HSSFFont font = workbook.createFont();
	    cell.setCellStyle(style);
	    font.setColor(color);
	    style.setFont(font);
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
	 	FileOutputStream outFile =new FileOutputStream(new File(excelpath));
		workbook.write(outFile);
		outFile.close();
	 	workbook.close();
		}
	
	public int getLastRowNum(int SheetNo ,String excelpath) throws InvalidFormatException, IOException{
		
		FileInputStream file = new FileInputStream(excelpath);
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sh = workbook.getSheetAt(SheetNo);
		int data=sh.getLastRowNum();
		workbook.close();	       
		return data;
		}
	
   
	
    
    
    public void singleCellMultirowComparison(String excelPath,int sheetNo,int rowNum,int colNum,ArrayList<String> al1,ArrayList<String> al2) throws InvalidFormatException, IOException{
		
		FileInputStream file = new FileInputStream(new File(excelPath));
		HSSFWorkbook workbook = new HSSFWorkbook(file);
		HSSFSheet sheet = workbook.getSheetAt(sheetNo);
		HSSFCell cell = null;
		FileOutputStream out;
		String[] data1 = null;
	    String[] data2 = null;
	    StringBuilder result= new StringBuilder();
	   for (int startIndex = 0,row=rowNum,col=colNum; startIndex<al1.size(); startIndex++,row++ ) {
	    	 data1 = al1.get(startIndex).split("\n");
	    	 data2 = al2.get(startIndex).split("\n");
	    	for(int i=0;i<data1.length;i++){
	    		// System.out.println("array :"+data1.length);
	    		//try{
	    		if(data1[i].equals(data2[i])){
	    			
	    			cell = sheet.getRow(row).getCell(col);
	    			result.append("PASS"+"\n");
	    		    String comparisonResult=result.toString();
	    		    cell.setCellValue(comparisonResult);
	    		   HSSFCellStyle style = workbook.createCellStyle();
	    		    	HSSFFont font = workbook.createFont();
		  			    style.setWrapText(true);
		  			    cell.setCellStyle(style);
		  			    font.setColor(HSSFColor.GREEN.index); 
		  		        style.setFont(font);
		  		        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
	    		         out = new FileOutputStream(excelPath);
				         workbook.write(out);
				         out.close();
	    			}
	    		   else{
	    			
	    			cell = sheet.getRow(row).getCell(col);
	    			result.append("FAIL"+"\n");
	    			String comparisonResult=result.toString();
	    			cell.setCellValue(comparisonResult);
	    			HSSFCellStyle style1 = workbook.createCellStyle();
	  			    HSSFFont font1 = workbook.createFont();
	  			    style1.setWrapText(true);
	  			    cell.setCellStyle(style1);
	  			    font1.setColor(HSSFColor.RED.index); 
	  		        style1.setFont(font1);
	  		        font1.setBoldweight(Font.BOLDWEIGHT_BOLD);
	    			out = new FileOutputStream(excelPath);
					workbook.write(out);
					out.close();
	  		  }
	    	
	    		/*catch(ArrayIndexOutOfBoundsException e){
	    			if (frmOpt == null) {
				        frmOpt = new JFrame();
				    }
				    frmOpt.setVisible(true);
				    frmOpt.setLocationRelativeTo(null);
				    frmOpt.setAlwaysOnTop(true);
					JOptionPane.showMessageDialog(frmOpt, "Total no. of slides in Salesforce and InputFile.xls are not matching");
	    			throw new SkipException("Total no. of slides in Salesforce and InputFile.xls are not matching");
	    			
	    		}*/
	    		
	    		
	    		
	    	}
	    	if(result.toString().contains("FAIL")){
				
    			HSSFCellStyle style1 = workbook.createCellStyle();
  			    HSSFFont font1 = workbook.createFont();
  			    style1.setWrapText(true);
  			    cell.setCellStyle(style1);
  			    font1.setColor(HSSFColor.RED.index); 
  		        style1.setFont(font1);
  		        font1.setBoldweight(Font.BOLDWEIGHT_BOLD);
  		         out = new FileOutputStream(excelPath);
				  workbook.write(out);
				  out.close();
    		
    		}
    		
	  
	  result.setLength(0);
	    }
	   workbook.close();
	}
    
 public List<String> allDifferentPass(String excelPath,int sheetNo,int rowNum,int colNum,ArrayList<String> al1) throws InvalidFormatException, IOException{
		
		FileInputStream file = new FileInputStream(new File(excelPath));
		HSSFWorkbook workbook = new HSSFWorkbook(file);
		HSSFSheet sheet = workbook.getSheetAt(sheetNo);
		HSSFCell cell = null;
		FileOutputStream out;
		String[] data = null;
		String result;
	    List<String> list = new LinkedList<String>();
	    for (int startIndex = 0,row=rowNum,col=colNum; startIndex<al1.size(); startIndex++,row++ ) {
	    	 data = al1.get(startIndex).split("\n");
	    	for(int i=0;i<data.length;i++){
	    		
	    		list.add(data[i]);
	    	
	    	}
	    	
		 	if(WebDriverCommonUtils.findDuplicates(list).size()==0){
		 		
		 		cell = sheet.getRow(row).getCell(col);
		 		result="PASS";
    		    cell.setCellValue(result);
    		    HSSFCellStyle style = workbook.createCellStyle();
    		    	HSSFFont font = workbook.createFont();
	  			    style.setWrapText(true);
	  			    cell.setCellStyle(style);
	  			    font.setColor(HSSFColor.GREEN.index); 
	  		        style.setFont(font);
	  		        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
    		         out = new FileOutputStream(excelPath);
			         workbook.write(out);
			         out.close();
			           
		 	}
		 	else{
		 		
    			cell = sheet.getRow(row).getCell(col);
    			result="FAIL";
    			cell.setCellValue(result);
    			HSSFCellStyle style1 = workbook.createCellStyle();
  			    HSSFFont font1 = workbook.createFont();
  			    style1.setWrapText(true);
  			    cell.setCellStyle(style1);
  			    font1.setColor(HSSFColor.RED.index); 
  		        style1.setFont(font1);
  		        font1.setBoldweight(Font.BOLDWEIGHT_BOLD);
    			out = new FileOutputStream(excelPath);
				workbook.write(out);
				out.close();
				  
		 	} 
		 	while (!list.isEmpty()) {
		        list.remove(0);
		    }
		       		
}
	    workbook.close();
	   return list;
	   }


	
}

