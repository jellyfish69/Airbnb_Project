
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.util.Scanner;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;

public class Main {
	@SuppressWarnings("deprecation")
	public static void main( String [] args ) {
	    try {
	    	
	    	int year_var;
	    	
	    	//Create output Excel file
	    	String filename = "/Users/rafaelchris/Desktop/NewExcelFile.xls" ;
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet o_sheet = workbook.createSheet("FirstSheet");  

            HSSFRow rowhead = o_sheet.createRow((short)0);
            rowhead.createCell(0).setCellValue("Year");
            rowhead.createCell(1).setCellValue("Month");
            rowhead.createCell(2).setCellValue("Day");
            rowhead.createCell(3).setCellValue("Clicks");
            rowhead.createCell(4).setCellValue("Name");

            
            int out_row_num=1;
            int year_temp = 2013;
            
	    	for(int j=2013; j<=2017; j++)
	    	{
	    		
	    		File dir = new File("/Users/rafaelchris/Desktop/airbnb");
	    	
	    		File[] files = dir.listFiles(new FilenameFilter() {
	    			@Override
	    			public boolean accept(File dir, String name) {
	    				return !name.equals(".DS_Store");
	    			}
	    		});
	    	
	    		//Opening the input Excel files
	    		for ( File file : files ) {
	    		
	    			POIFSFileSystem fs = new POIFSFileSystem( file );
	    			HSSFWorkbook wb = new HSSFWorkbook(fs);

	    			for (int i = 0; i < wb.getNumberOfSheets(); i++) {
	    				HSSFSheet sheet = wb.getSheetAt(i);

	    				//position of cell containing years inside in_sheet 2,0  
	    				//position of cell containing name inside in_sheet 1,1 ~ 1,2 
	 	            
	    				HSSFRow in_row = sheet.getRow(2);
	    				HSSFRow checker_row = sheet.getRow(3);
	    				HSSFRow name_row = sheet.getRow(1);
	 	           
	 	            
	    				String boo = in_row.getCell(0).toString();
	    				boo = StringUtils.left(boo,4);
	    				year_var = Integer.valueOf(boo);
	 	            
	    				//cell for year
	    				Cell c0 = in_row.getCell(0);
	    				//cell for month
	    				Cell c1 = in_row.getCell(1);
	 	        
	    				//cell for day
	    				Cell c2 = in_row.getCell(2);
	    				//cell for the next day, indicating the change of the month 
	    				Cell checker = in_row.getCell(2);
	 	        
	    				//cell for clicks
	    				Cell c3 = in_row.getCell(4);
	    				//cell for name
	    				//the name will may be inside one or either two cells
	    				Cell c4 = name_row.getCell(1);
	    				Cell c5 = name_row.getCell(2);
	 	
	    				HSSFRow out_row = o_sheet.createRow(out_row_num);
	 	        
	    				if(year_var==j)
	    				{
	    					int in_row_num=2;
	    					double sum = 0;
	 	        	
	    					try {
	 	        		
	    						if(c0.getNumericCellValue()!=year_temp)
	    						{
	    							sum = 0;
	    							year_temp=j;  
	    						}
	 	        		
	    					//loop responsible for each current in_sheet
	    					while(in_row!=null) 
	    					{
	    						double num = 0;
	 	        		  
	    						//Inspect whether the reading of the in_sheet has come to an end 
	    						//checker is either pointing to the next row containing legitimate values or null
	    						checker_row = sheet.getRow(in_row_num + 1); 
	    						if (checker_row!=null)
	    							checker = checker_row.getCell(2);
	    						else 
	    						{
	    							checker_row = sheet.getRow(in_row_num);
	    							checker = checker_row.getCell(2);
	    						}
	 	        		  
	    						sum+=c3.getNumericCellValue();
	 	        		 
	    						out_row.createCell(0).setCellValue(c0.getNumericCellValue());
	    						out_row.createCell(1).setCellValue(c1.getNumericCellValue());
	    						out_row.createCell(2).setCellValue(c2.getNumericCellValue());
	 	        		 
	    						int length = String.valueOf(c2.getNumericCellValue()).length()-2;
	 		 	       
	    						double sum_temp = 0;
	    						int flag = 0;
	 	        		
	    						//if the current day of the month is a two-digit number 
	    						if(length==2)
	    						{
	    							out_row.createCell(3).setCellValue(in_row.getCell(4).getNumericCellValue());
	    							out_row.createCell(4).setCellValue(c4.toString());
	 	    			
	    							if(c5!=null)
	    								out_row.createCell(4).setCellValue(c5.toString());
	 	    			
	    							//check if checker is pointing to a single-digit number
	    							//each number is represented as e.g. 3.0 , 13.0
	    							if(String.valueOf(checker.getNumericCellValue()).length() == 3)
	    							{
	    								//months with 31 days
	    								if(c1.getNumericCellValue()==1.0 || c1.getNumericCellValue()==3.0 || c1.getNumericCellValue()==5.0 || + 
	    										+ c1.getNumericCellValue()==7.0 || c1.getNumericCellValue()==8.0 || c1.getNumericCellValue()==10.0 || + 
	    											+ c1.getNumericCellValue()==12.0)
	    								{
	    									flag=1;
	    									c3 = in_row.getCell(4);
	    									num =((31-c2.getNumericCellValue())*c3.getNumericCellValue())/7;
	 		 	    			  
	    								}
	    								
	    								//months with 30 days
	    								if(c1.getNumericCellValue()==4.0 || c1.getNumericCellValue()==6.0 || c1.getNumericCellValue()==9.0 || + 
	    									+ c1.getNumericCellValue()==11.0)
	    								{
	    									flag = 1;
	    									c3 = in_row.getCell(4);
	    									num =((30-c2.getNumericCellValue())*c3.getNumericCellValue())/7;
	    								}
	 		 	    		
	    								//February
	    								if(c1.getNumericCellValue()==2.0 && flag==0)
	    								{
	    									if(c0.getNumericCellValue()%4==0 && c0.getNumericCellValue()%100!=0)
	    									{
	    										//leap year
	    										c3 = in_row.getCell(4);
	    										num =((29-c2.getNumericCellValue())*c3.getNumericCellValue())/7;
	    									}
	    									else if(c0.getNumericCellValue()%400==0)
	    									{
	    										//leap year
	    										c3 = in_row.getCell(4);
	    										num =((29-c2.getNumericCellValue())*c3.getNumericCellValue())/7;
	    									}
	    									else 
	    									{
	    										//normal year
	    										c3 = in_row.getCell(4);
	    										num =((28-c2.getNumericCellValue())*c3.getNumericCellValue())/7;
	    									}
	    								}
		 		 	    	
		 		 	    	
	    							out_row.createCell(3).setCellValue(sum+num);
	    							out_row.createCell(4).setCellValue(c4.toString());
	    							if(c5!=null)
	    								out_row.createCell(4).setCellValue(c5.toString());
		 	    			 
		 	    		
	    							sum_temp = c3.getNumericCellValue();
	    							out_row_num++;
	    							in_row_num++;
	    							in_row = sheet.getRow(in_row_num);
	    							c0 = in_row.getCell(0);
	    							c1 = in_row.getCell(1);
	    							c2 = in_row.getCell(2);
	    							c3 = in_row.getCell(4);
	    							
	    							//filling the next row, each time there is a change of month

	    							out_row = o_sheet.createRow(out_row_num);
		 	        	  
	    							out_row.createCell(0).setCellValue(c0.getNumericCellValue());
	    							out_row.createCell(1).setCellValue(c1.getNumericCellValue());
	    							out_row.createCell(2).setCellValue(c2.getNumericCellValue());
	    							out_row.createCell(3).setCellValue(sum_temp-num);
	    							out_row.createCell(4).setCellValue(c4.toString());
	    							if(c5!=null)
	    								out_row.createCell(4).setCellValue(c5.toString());
	 	    			 
	    							out_row_num++;
	    							in_row_num++;
	    							in_row = sheet.getRow(in_row_num);
	    							c0 = in_row.getCell(0);
	    							c1 = in_row.getCell(1);
	    							c2 = in_row.getCell(2);
	    							c3 = in_row.getCell(4);
		 	        	  	
		 	        	  	
	    							out_row = o_sheet.createRow(out_row_num);
		 	        	
	    							sum = sum_temp - num;
	 		 	    		 
	    							} else
	    								{
	    									//if day is a two-digit number but checker is a single-digit one
	    									out_row.createCell(3).setCellValue(c3.getNumericCellValue());
	    									out_row.createCell(4).setCellValue(c4.toString());
	    									if(c5!=null)
	    										out_row.createCell(4).setCellValue(c5.toString());
		 		 	    	
	    									out_row_num++;
	    									in_row_num++;
	    									in_row = sheet.getRow(in_row_num);
	    									c0 = in_row.getCell(0);
	    									c1 = in_row.getCell(1);
	    									c2 = in_row.getCell(2);
	    									c3 = in_row.getCell(4);
	    									
	    									out_row = o_sheet.createRow(out_row_num);
	    								}
	 		 	    	
	    						}
	    						else 
	    						{
	    							//if day is a two-digit number
	    							out_row.createCell(3).setCellValue(c3.getNumericCellValue());
	    							out_row.createCell(4).setCellValue(c4.toString());
	    							if(c5!=null)
	    								out_row.createCell(4).setCellValue(c5.toString());
	 		 	    	   
	    							out_row_num++;
	    							in_row_num++;
	    							in_row = sheet.getRow(in_row_num);
	    							c0 = in_row.getCell(0);
	    							c1 = in_row.getCell(1);
	    							c2 = in_row.getCell(2);
	    							c3 = in_row.getCell(4);
	    							
	    							out_row = o_sheet.createRow(out_row_num);
	    						}
	    					}
	    					} catch(NullPointerException NPE)
	    					{
	    						break;
	    					}
	    				}
	    			else break;
	    			}
	    		}
	    	}
	    	
	    	//Closing the output Excel file
	    	 FileOutputStream fileOut = new FileOutputStream(filename);
	         workbook.write(fileOut);
	         fileOut.close();
	         workbook.close();
	         System.out.println("Your excel file has been generated!");
	    }
	    
	    catch ( IOException ex ) {
	        ex.printStackTrace();
	    }   
	}
}
