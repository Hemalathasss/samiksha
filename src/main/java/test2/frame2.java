package test2;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Date;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class frame2 {
	public static void main(String[] args)throws IOException 

{
	//Create an object of File class to open xlsx file

    File file =    new File("C:\\Users\\91994\\eclipse-workspace\\frame2\\excel\\data.xlsx");

    //Create an object of FileInputStream class to read excel file

    FileInputStream inputStream = new FileInputStream(file);
    
   Workbook w = new XSSFWorkbook(inputStream);
    Sheet s=w.getSheet("Sheet1");
    Row r =s.getRow(0);
   Cell c=r.getCell(0);
//   System.out.println(c); 
   for(int i = 0;i <s.getPhysicalNumberOfRows();i++)
   {
	   Row r1 = s.getRow(i);
       for(int j = 0;j<r.getPhysicalNumberOfCells();j++)
       {
    	   
    	   Cell cell=r1.getCell(j);
    	   int CType =cell.getCellType();
          // System.out.println(CType);
           if(CType==1) {
        	   String data=cell.getStringCellValue();
        	   System.out.println(data);
           }
          if(CType==0) 
          {
        	  if(DateUtil.isCellDateFormatted(cell))
        	  {
        	  Date CV=(Date) cell.getDateCellValue();
        	  SimpleDateFormat Df=new SimpleDateFormat("MM/dd/yyyy");
        	  String data=Df.format(CV);
        	  System.out.println(data);
          }
        	  else
        	  {
        		  double db=cell.getNumericCellValue();
        		  long l= (long)db;
        		  String data=String.valueOf(l);
        		  System.out.println(data);
        	  }
        	
       }
          
    }
    }
  
}
}


