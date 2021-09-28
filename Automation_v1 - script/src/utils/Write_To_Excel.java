package utils;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;


public class Write_To_Excel {
	
	private String fileName;
	private FileOutputStream fileOut;
	private HSSFWorkbook workbook;
	private HSSFSheet sheet;
	private static final Integer START_TIMEOUT = 3;
	private static final Integer START_COLLATERAL = 6;
	private static final Integer START_OUTPUTTEMP = 22;
	
	public Write_To_Excel(String fileName, String sheetName)
	{
		try {
            this.fileName = fileName;
            this.workbook = new HSSFWorkbook();
            this.sheet = workbook.createSheet(sheetName);  
            ArrayList<String> columnHeader = new ArrayList<String>();
            HSSFRow rowhead = sheet.createRow((short)0);
            Iterator<String> iter;
    		Integer ctr = 0;
            
            //STEP
            columnHeader.add("Step Name");
            columnHeader.add("Step Default Channel Class");
            columnHeader.add("Step Default Channel Instance");
            //TIMEOUT
            columnHeader.add("Timeout Days");
            columnHeader.add("Action Type");
            columnHeader.add("Action Connection");
            //COLLATERAL
            columnHeader.add("Collateral Name");
            columnHeader.add("Collateral Type");
            columnHeader.add("Collateral Email");
            columnHeader.add("DMC Message ID");
            columnHeader.add("Subject Line");
            columnHeader.add("Incentive");
            //AB TESTING
            columnHeader.add("Placeholder");
            columnHeader.add("Required Response %");
            columnHeader.add("Response Type");
            columnHeader.add("Min. Wait Period");
            columnHeader.add("Max. Wait Period");
            columnHeader.add("Candidate 1");
            columnHeader.add("Candidate 2");
            columnHeader.add("Candidate 3");
            columnHeader.add("Candidate 4");
            columnHeader.add("Candidate 5");
            
            //OUTPUT TEMPLATE
            columnHeader.add("Output Name");
            columnHeader.add("Output Type");
            columnHeader.add("Output Delivery Setting");
            columnHeader.add("Output Domain ID");
            columnHeader.add("Output Speed");
            columnHeader.add("Email Reply Handling");
            columnHeader.add("Email Friendly name");
            columnHeader.add("Email From name");
            columnHeader.add("Sending To");
            columnHeader.add("Email Address Field");
            columnHeader.add("# of days after processing");
            columnHeader.add("Send Message Time");
            
            iter = columnHeader.iterator();
		      while (iter.hasNext()) 
		      {
		    	  Cell cell = rowhead.createCell(0+ctr);

		    	  
		    	  if(ctr == 11)
		    	  {  
		    		  this.sheet.setColumnWidth(ctr, 10240);
		    	  }
		    	  cell.setCellValue(iter.next());
		    	  ctr++;
		      }
    	      
            this.fileOut = new FileOutputStream(fileName);
            this.workbook.write(this.fileOut);


        } catch ( Exception ex ) {
            System.out.println(ex);
        }
		
	}
	

	
	public void writeStepProperties(Integer rowNum, ArrayList<String> prop)
	{
		HSSFRow stepRow = this.sheet.createRow(rowNum);
		
		Iterator<String> iter = prop.iterator();
		Integer ctr = 0;
	      while (iter.hasNext()) 
	      {
	    	  Cell cell = stepRow.createCell(0+ctr);
	    	  cell.setCellValue(iter.next());
	    	  ctr++;
	      }
		
		try {
			this.fileOut = new FileOutputStream(this.fileName);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        try {
			this.workbook.write(this.fileOut);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public void writeTimeoutProperties(Integer rowNum, ArrayList<String> prop)
	{
		HSSFRow rowTimeout;
		
		if (rowNum > this.sheet.getLastRowNum() )
		{
			rowTimeout = this.sheet.createRow(rowNum);
		}
		else
		{
			rowTimeout = this.sheet.getRow(rowNum);
		}
		
		
		Iterator<String> iter = prop.iterator();
		Integer ctr = 0;
	      while (iter.hasNext()) 
	      {
	    	  Cell cell = rowTimeout.createCell(START_TIMEOUT+ctr);

	    	  cell.setCellValue(iter.next());
	    	  ctr++;
	      }
		
		try {
			this.fileOut = new FileOutputStream(this.fileName);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        try {
			this.workbook.write(this.fileOut);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public void writeCollateralProperties(Integer rowNum, ArrayList<String> prop)
	{
		HSSFRow rowCollateral;
		
		if (rowNum > this.sheet.getLastRowNum() )
		{
			rowCollateral = this.sheet.createRow(rowNum);
		}
		else
		{
			rowCollateral = this.sheet.getRow(rowNum);
		}
		
		
		Iterator<String> iter = prop.iterator();
		Integer ctr = 0;
	      while (iter.hasNext()) 
	      {
	    	  Cell cell = rowCollateral.createCell(START_COLLATERAL+ctr);
	    	  cell.setCellValue(iter.next());
	    	  ctr++;
	      }
		
		try {
			this.fileOut = new FileOutputStream(this.fileName);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        try {
			this.workbook.write(this.fileOut);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public void writeOutputTempProperties(Integer rowNum, ArrayList<String> prop)
	{
		HSSFRow rowOutputTemp;
		
		if (rowNum > this.sheet.getLastRowNum() )
		{
			rowOutputTemp = this.sheet.createRow(rowNum);
		}
		else
		{
			rowOutputTemp = this.sheet.getRow(rowNum);
		}
		
		
		Iterator<String> iter = prop.iterator();
		Integer ctr = 0;
	      while (iter.hasNext()) 
	      {
	    	  Cell cell = rowOutputTemp.createCell(START_OUTPUTTEMP+ctr);

	    	  cell.setCellValue(iter.next());
	    	  ctr++;
	      }
		
		try {
			this.fileOut = new FileOutputStream(this.fileName);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        try {
			this.workbook.write(this.fileOut);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public void closeExcelFile()
	{
		try {
			this.fileOut.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	
	public void setAutoColWidth()
	{
		HSSFRow row = this.sheet.getRow(0);
		Iterator<Cell> cellIterator = row.cellIterator();
		CellStyle style = workbook.createCellStyle();
	    Font font = workbook.createFont();
	    font.setBold(true);
	    style.setFont(font);
		
		while (cellIterator.hasNext()) 
		{
            Cell cell = cellIterator.next();
            int columnIndex = cell.getColumnIndex();
            
            cell.setCellStyle(style);
            
            if(columnIndex != 11)
            {
            	this.sheet.autoSizeColumn(columnIndex);
            }
            
        }
		
		try {
			this.fileOut = new FileOutputStream(this.fileName);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        try {
			this.workbook.write(this.fileOut);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
