package com.exelatech.linkchecker;

import java.io.FileInputStream;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ExcelReader {

			
	//*******************Initialization*****************************************************************************
		   public static XSSFSheet ExcelWSheet;
		   public static XSSFWorkbook ExcelWBook;
		   public static XSSFCell CellObj;
		   public static XSSFRow Row;
		   public static FileInputStream ExcelFile;
		   public static int Datarow=0, datarownum,rowstart = 0,rowends = 0;
		   public static String[] ScenarioNo;
		   public static int MessageToBeValidateColNum,ValidationTypeColNum,ToBeValidateColNum,TestDataRowNum;

	//**************************************************************************************************************   
		   
	/**
	* <p>This method is to set the File path and to open the Excel file</p>
	* @param SheetName
	* @return int
	* @throws Exception
	* @author Pankaj.Barlota
	*/
		 
	public static int setExcelFile(String SheetName) throws Exception 
	{
		ExcelFile = new FileInputStream(SheetName);
		ExcelWBook = new XSSFWorkbook(ExcelFile);
		ExcelWSheet = ExcelWBook.getSheet("QueryData");
		//System.out.println("Current Sheet Name is : "+SheetName);
		int rowCount = ExcelWSheet.getLastRowNum()-ExcelWSheet.getFirstRowNum();
		//System.out.println("Total number of rows in modules sheet is : "+rowCount);
		return rowCount;
	}
	
	
	 //**************************************************************************************************************	            
	/**   
	* <p>This method is used to get the data from specific cell</p>
	* @param RowNum
	* @param ColNum
	* @return Object
	* @author Pankaj.Barlota
	*/
	public static Object getCellData(int RowNum, int ColNum)
	{
		XSSFCell cell=ExcelWSheet.getRow(RowNum).getCell(ColNum);
		switch (cell.getCellType()) {
			  case Cell.CELL_TYPE_STRING:
			       return cell.getStringCellValue();
			             
			  case Cell.CELL_TYPE_BOOLEAN:
			       return cell.getBooleanCellValue();
			             
			  case Cell.CELL_TYPE_NUMERIC:
			       DateFormat df = new SimpleDateFormat("MM/dd/yyyy");
			       if (HSSFDateUtil.isCellDateFormatted(cell))
				          return    df.format(cell.getDateCellValue());
			       else return new BigDecimal(cell.getNumericCellValue()).toPlainString();
			                    
			  case Cell.CELL_TYPE_ERROR :
			       return cell.getStringCellValue();
			                
		} 
		return null;
	}	       

	//**************************************************************************************************************
	/**
	* <p>This method is used to get the column number</p>
	* @param row
	* @param ColumnName
	* @return int
	* @author Pankaj.Barlota
	*/
	      
	public static int getColumnNumber(int row,String ColumnName)
	{
		XSSFRow rowheader=ExcelWSheet.getRow(row);
		        	
		int columnnumber;
		ForLoop:
		for (columnnumber=rowheader.getFirstCellNum();columnnumber<rowheader.getLastCellNum();columnnumber++){
		     XSSFCell cellobj=rowheader.getCell(columnnumber);
		     String header=cellobj.getStringCellValue();
		     if (header.equalsIgnoreCase(ColumnName)) break ForLoop;      			
		        		        		
		}
		return columnnumber;
	}	
	

}
