package com.generic.utilities;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CsvHandler {

	private static final Logger LOGGER = Logg.createLogger();
	public CsvHandler() {
    }
    
    /**
     * This Method is to convert CSV file to Excel file format(XLSX)
     * Excel file will get created at path specified in parameter xlsxFilePath
     * @param  String csvFilePath 
     * @param String xlsxFilePath 
     * @return void
     * @throws IOException
     */
    public void convertToXLSX(String csvFilePath,String xlsxFilePath) throws IOException ,FileNotFoundException
    {
    	 
    	LOGGER.info("****************Converting CSV File to XLSX File *****************************");
    	LOGGER.info("Verifying CSV Filename exist or not");
    	FileOutputStream fileOutputStream=null;
    	BufferedReader br =null;
    	if(csvFilePath.toLowerCase().endsWith(".csv"))
    	{
    	try{
    		File fileName=new File(csvFilePath); 
    		if(!fileName.exists() || !fileName.isFile())
    		{
    			throw new FileNotFoundException("File does not exist at location"+csvFilePath+"\n Check you have provided valid file name");
    		}
    		else
    		{
    			LOGGER.info("Verified that given csv file path is valid.");   		
    			XSSFWorkbook wb = new XSSFWorkbook();
    	
    			XSSFSheet sheet = wb.createSheet(fileName.getName().replace(".csv", ""));
    			String currentLine=null;
    			int RowNum=0;
 
    			br= new BufferedReader(new FileReader(fileName));
    			int x=1;
    			while ((currentLine = br.readLine()) != null) 
    			{
    				String str[] = currentLine.split(",");
    				LOGGER.info("Creating Row "+(x++));
    				XSSFRow currentRow=sheet.createRow(RowNum);
    				for(int i=0;i<str.length;i++){
    					currentRow.createCell(i).setCellValue(str[i]);
    				}
    				RowNum++;
    			}
    		
    			File xlsxFile=new File(xlsxFilePath);
    			LOGGER.info("Verifying Given XLSX file Exist or not");
    			if(xlsxFile.exists())
    			{
    				LOGGER.info("XLSX File is Exist");
    				if(xlsxFile.isFile())
    				{		
    					LOGGER.info("Given File is of Type FILE");
    					fileOutputStream = new FileOutputStream(xlsxFile);
    				}
    				else if(xlsxFile.isDirectory())
    				{
    					LOGGER.info("Given File is of TYPE Directory");
    					System.out.println("Provided path is directory so creating new XLSX File with name of CSV file");
    					String newXlsxFilePath=xlsxFile.getAbsolutePath().concat("\\").concat(fileName.getName().replace(".csv", ".xlsx"));
    					LOGGER.info("Creating file in Directory with same name as provided for CSV");
    					fileOutputStream =new FileOutputStream(newXlsxFilePath);
    				}
    			}
    			else
    			{
    				fileOutputStream = new FileOutputStream(xlsxFilePath);
    			}
    			LOGGER.info("Writing file to Excel sheet");
    			wb.write(fileOutputStream);
    			LOGGER.info("CSV file has been converted to Excel Succesfully at location"+xlsxFilePath); 
    			LOGGER.info("Closed the File InputStream");	
    			}
    		}
    		catch(FileNotFoundException e)
    		{
    			LOGGER.info(e.getMessage());
    		}
    		catch (IOException e) 
    		{
    	
    			LOGGER.info(e.getMessage());
    		}
    		finally{
    			if(br!=null || fileOutputStream != null){
    			LOGGER.info("Closing Buffered reader and FileInputStream ");
    			fileOutputStream.close();
    			br.close();
    			}
    		}
    		}
    	else
		{
			LOGGER.error("Not Valid CSV file : "+csvFilePath);
			throw new FileNotFoundException("Not Valid CSV File :"+csvFilePath);    		 
		}   	
  }
   
    /**
     * This Method is to convert CSV file to XLSX file Provided only CSv File path.
     * It will Create XLSX file at the location where CSV file is resides.
     * @param String csvFilePath 
     * @return void
     * @throws IOException ,FileNotFoundException
     */
    
  public void convertToXLSX(String csvFilePath) throws IOException ,FileNotFoundException
  {
	  LOGGER.info("*****************Convert CSV to Excel (XLSX) ******************************* ");
	  LOGGER.info("Provided only csvFile path. It will create excel file in same folder the csv file resides");
	  
	  //Check whether given file path contains csv file
	  
	  if(csvFilePath.toLowerCase().endsWith(".csv"))
	  {
		  try{
			  File csvFile=new File(csvFilePath);
			  if(csvFile.exists() && csvFile.isFile())
			  {
				  String csvFileFolderPath=csvFile.getParentFile().getAbsolutePath().toString().concat("\\");
				  String xlsxFilePath=csvFileFolderPath.concat(csvFile.getName().replace(".csv", ".xlsx"));
				  convertToXLSX(csvFilePath, xlsxFilePath);
			  }			  
			  else
			  {
				  LOGGER.error("CSV File does not exist at location "+csvFilePath);
				  throw new FileNotFoundException("File does not exist at location"+csvFilePath);
			  }
		  	}
		  catch(FileNotFoundException ex)
		  {
			  LOGGER.info(ex.getMessage());
		  }
	  }
	  else
	  {
		  LOGGER.error("Not valid CSV file at "+csvFilePath);
		  throw new FileNotFoundException("Not valid CSV file at "+csvFilePath);
	  }
	  
  }
  
  /**
   * This Method will convert CSV file into Excel(97-2003) Format.
   * Excel file will get created at path specified in parameter xlsFilePath
   * @param  String csvFilePath - File path for CSV 
   * @param String xlsFilePath - File path for XLS
   * @return void  
   * @throws IOException 	
   */
  
  public void convertToXLS_97_2003_format(String csvFilePath,String xlsFilePath) throws IOException ,FileNotFoundException
  {
	LOGGER.info("****************Converting CSV File to XLS(Excel 97-2003) File *****************************");
  	LOGGER.info("Verifying CSV Filename exist or not");
  	FileOutputStream fileOutputStream=null;
  	BufferedReader br =null;
  	if(csvFilePath.toLowerCase().endsWith(".csv"))
	{
  	 try{
  		 File fileName=new File(csvFilePath); 
  		 if(!fileName.exists() || !fileName.isFile())
  		 {
  			throw new FileNotFoundException("File does not exist at location"+csvFilePath+"\n Check you have provided valid file name");
  		 }
  		 else
  		 {
  			LOGGER.info("Verified that given csv file path is valid.");   		
  			HSSFWorkbook wb = new HSSFWorkbook();
  	
  			HSSFSheet sheet = wb.createSheet(fileName.getName().replace(".csv", ""));
  			String currentLine=null;
  			int RowNum=0;

  			br= new BufferedReader(new FileReader(fileName));
  			int x=1;
  			while ((currentLine = br.readLine()) != null) 
  			{
  				String str[] = currentLine.split(",");
  				LOGGER.info("Creating Row "+(x++));
  				HSSFRow currentRow=sheet.createRow(RowNum);
  				for(int i=0;i<str.length;i++){
  					currentRow.createCell(i).setCellValue(str[i]);
  				}
  				RowNum++;
  			}
  		
  			File xlsFile=new File(xlsFilePath);
  			LOGGER.info("Verifying Given XLS(Excel 97-2003) file Exist or not");
  			if(xlsFile.exists())
  			{
  				LOGGER.info("XLS File is Exist");
  				if(xlsFile.isFile())
  				{		
  					LOGGER.info("Given File is of Type FILE");
  					fileOutputStream = new FileOutputStream(xlsFile);
  				}
  				else if(xlsFile.isDirectory())
  				{
  					LOGGER.info("Given File is of TYPE Directory");
  					System.out.println("Provided path is a directory.Creating new XLS (Excel 97-2003) File with name of CSV file");
  					String newXlsFilePath=xlsFile.getAbsolutePath().concat("\\").concat(fileName.getName().replace(".csv", ".xls"));
  					LOGGER.info("Creating file in Directory with same name as provided for CSV");
  					fileOutputStream =new FileOutputStream(newXlsFilePath);
  				}
  			}
  			else
  			{
  				fileOutputStream = new FileOutputStream(xlsFilePath);
  			}
  			LOGGER.info("Writing file to Excel sheet");
  			wb.write(fileOutputStream);
  			LOGGER.info("CSV file has been converted to Excel Succesfully at location"+xlsFilePath); 
  			LOGGER.info("Closed the File InputStream");
  		 }	
  		}
  		catch(FileNotFoundException e)
  		{
  			LOGGER.error(e.getMessage());
  		}
  		catch (IOException e) 
  		{
  			LOGGER.error(e.getMessage());
  		}
  		finally
  		{
  			if(br!=null || fileOutputStream != null)
  			{
  				LOGGER.info("Closing Buffered reader and FileInputStream ");
  				fileOutputStream.close();
  				br.close();
  			}
  		}
	}
  	else
  	{
  		LOGGER.error("Not Valid CSV file : "+csvFilePath);
		throw new FileNotFoundException("Not Valid CSV File :"+csvFilePath);    		 
	}   
}
  
  /**
   * This method will convert CSV file into Excel(97-2003) Format. 
   * Excel file will get created at the same location where CSV file is present.
   * @param String csvFilePath
   * @return Void  
   */
  public void convertToXLS_97_2003_Format(String csvFilePath) throws IOException ,FileNotFoundException
  {
	  LOGGER.info("*****************Convert CSV to Excel (97-2003) ******************************* ");
	  LOGGER.info("Provided only csvFile path. It will create excel Excel (97-2003) file in same directory the csv file resides");
	  if(csvFilePath.endsWith(".csv"))
	  {
		  try
		  {
			  File csvFile=new File(csvFilePath);
			  if(csvFile.exists() && csvFile.isFile())
			  {	
				  String csvFileFolderPath=csvFile.getParentFile().getAbsolutePath().toString().concat("\\");
				  String xlsFilePath=csvFileFolderPath.concat(csvFile.getName().replace(".csv", ".xls"));
				  convertToXLSX(csvFilePath, xlsFilePath);
			  }
			  else
			  {
				  LOGGER.error("Not valid CSV file at "+csvFilePath);
				  throw new FileNotFoundException("Not valid CSV file at "+csvFilePath);
			  }
		  }
		  catch(FileNotFoundException ex)
		  {
			  LOGGER.error(ex.getMessage());
		  }
	  }
	  else
	  {
		  LOGGER.error("CSV File does not exist at location "+csvFilePath);
		  throw new FileNotFoundException("File does not exist at location"+csvFilePath);
	  }	  
  }
  
  /**
   *  This method returns the Total number of rows in CSV file
   *  @param String  
   *  @return int
   * @throws IOException 
   */
  
  public int getTotalRows(String csvFilePath) throws IOException,FileNotFoundException
  {
	  int totalRows=0;
	  BufferedReader br=null;
	  if(csvFilePath.endsWith(".csv"))
	  {
		  try
		  {	  
			  File csvFile =new File(csvFilePath);
	  		if(csvFile.isFile() && csvFile.getName().endsWith(".csv"))
	  		{
	  			br = new BufferedReader(new FileReader(csvFilePath));
		       
	  				while ((csvFilePath = br.readLine()) != null) 
	  				{
	  					totalRows++;
	  				}
	  		}
	  		else
	  		{
	  			LOGGER.error("Not valid Csv File "+csvFilePath);
	  			throw new FileNotFoundException("Not valid Csv File "+csvFilePath);
	  		}
		  }
		  catch (FileNotFoundException e) 
		  {
			  LOGGER.error("File not Found at location "+csvFilePath);
			  LOGGER.error(e.getMessage());
		  }
		  catch (IOException e) 
		  {
			  LOGGER.error("Failed to read/ write File");
			  LOGGER.error(e.getMessage());
		  }
		  finally
		  {
			  if(br!=null){
				  br.close();}
		  }
	  }
	  else
	  {
		  LOGGER.error("Not valid CSV file path"+csvFilePath);
		  throw new FileNotFoundException("Not valid CSV file path at "+csvFilePath);
	  }	 
	  return totalRows;
  }
  
  /**
   * This method returns specific row specified by index
   * @param String csvFilepath 
   * 		int index   
   * 
   * @return String[] 
   */
  
  public String[] getRow(String csvFilePath,int index) throws IOException
  {
	  LOGGER.info("********************* In method getRow ny index *********************");
	  int rowNumber=0;
	  String row[]=null;
	  String lineToRead=null;
	  BufferedReader br =null;
	  LOGGER.info("Checking index is neither null nor greater than total number of rows");
	  try{  
	  if(index > getTotalRows(csvFilePath) || index <= 0)
	  {
		  LOGGER.error("Not valid index value"+index);
		  throw new ArrayIndexOutOfBoundsException("Index value mismatched error. Value = "+index);	  
	  }
	  else
	  {
		  File csvFile =new File(csvFilePath);
		  if(csvFile.exists())
		  {	
			  if(csvFile.isFile() && csvFile.getName().endsWith(".csv"))
			  {
				  br = new BufferedReader(new FileReader(csvFilePath));
		       
				  while ((lineToRead = br.readLine()) != null) 
				  {
					  if(rowNumber==index-1)
					  {
						  row=lineToRead.split(",");
					  }
					  rowNumber++;
				  }
				  br.close();
			  }
		  }
		  else
		  {
			  LOGGER.error("File does not Exist at location : "+csvFilePath);
			  throw new FileNotFoundException("Not valid CSV file present at location : "+csvFilePath); 
			   
		  }	
	  }
	  }
	  catch(ArrayIndexOutOfBoundsException e)
	  {
		  LOGGER.error(e.getMessage());
	  }
	  catch(FileNotFoundException e)
	  {
		  LOGGER.error(e.getMessage());
	  }
	  catch(IOException e)
	  {
		  LOGGER.error(e.getMessage());
	  }
	  finally
	  {
		  if(br!=null)
		  {
			  br.close();
		  }
	  }
	  return row;
  }
  
  /**
   * This method returns 1st row (Header) of CSV file 
   * @param String csvFilepath 
   * @return String[]
   */
  
  public String[] getHeader(String csvFilePath) throws IOException
  {
	  LOGGER.info("Returning 1st row (Header) ");
	  return getRow(csvFilePath, 1);
  }
  
  
  /**
   * This method returns last row of CSV file 
   * 
   * @param String csvFilepath
   * @return String[]
   */ 
  
  public String[] getLastRow(String csvFilePath) throws IOException
  {
	  return getRow(csvFilePath, getTotalRows(csvFilePath));
  }
  
 
  
  /**
   * This method returns specific Column specified by index 
   * Column is nothing but an array of string.where string represents column cell
   * @param String csvFilepath
   * 		int index 
   * @return String[]
   */
  public String[] getColumn(String csvFilePath,int index) throws IOException
  {
	  String columnValues[]=null;
	  BufferedReader br=null;
	  LOGGER.info("********************** In Method getColumn by index ********************");
	  LOGGER.debug("Checking index Value is not negative");
	  try
	  {
		  if(index <= 0)
		  {
			  LOGGER.error("Index Contain Negative Value");
			  throw new ArrayIndexOutOfBoundsException("Not valid index");
		  }
		  else
		  {
			  columnValues=new String[getTotalRows(csvFilePath)];
			  String row[]=null;
			  int column=0;
			  String lineToRead=null;
			  br = new BufferedReader(new FileReader(csvFilePath));
			  while ((lineToRead = br.readLine()) != null) 
			  {	
	      			row=lineToRead.split(",");
	      			LOGGER.debug("Copying coulumn value at index"+index+" to array");
	      			columnValues[column++]=row[index-1];
	      		}
	      }
	  }
	  catch(ArrayIndexOutOfBoundsException e)
	  {
		LOGGER.error("getColumn by Index error "+e.getMessage());  
	  }
	  finally
	  {
		  if(br!=null)
		  {
			  LOGGER.info("Closing BufferedReader..." );
			  br.close();
		  }
	  }
	  		LOGGER.info("Returning Column values");
	      	return columnValues;
	  }
  
	  
  /**
   * This method returns specific row specified by name 
   * @param String csvFilepath
   * 		String columnName 
   * @return String[]
   * @throws IOException
   */
  public String[] getColumn(String csvFilePath,String columnName) throws IOException ,FileNotFoundException
  {
	  String columnValues[]=new String[getTotalRows(csvFilePath)];
	  LOGGER.info("**************************** In method getColumn by Name *****************************");
	  LOGGER.info("Cheking gievn file path contains valid CSV file");
	  if(!csvFilePath.endsWith(".csv"))
	  {
		  LOGGER.error("Not valid CSV file...");
		  throw new FileNotFoundException("Not valid CSV File");
	  }
	  else
	  {
		  LOGGER.info("Verified filepath Contains valid CSV file");
		  String listOfColumnHeaders[]=getHeader(csvFilePath);
		  int columnIndex=1;
		  boolean found=false;
		  for(String column:listOfColumnHeaders)
		  {
			  LOGGER.debug("Verifying "+columnIndex+"Coulmn Contains given Column Name.");
			  if(column.equalsIgnoreCase(columnName)){
				  found=true;
				  LOGGER.debug("Given Column Name found at Index"+columnIndex);
				  LOGGER.debug("Fetching Column values At index"+columnIndex);
				  columnValues=getColumn(csvFilePath, columnIndex);
				  break;
			  }
			  else
			  {
				  columnIndex++;
			  }		  
		  }
		  if(found==false)
		  {
			  LOGGER.error("Coulmn name '"+columnName+"' doest not exist");
		  }
	  }
	  LOGGER.info("Returning Column Values...");
	      	return columnValues;
	  }

}   

