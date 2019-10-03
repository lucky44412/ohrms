package datadriven;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class creatingexcelTest {
  @Test
  public void f() throws IOException {
	String filepath="E:\\seleniumworkspace2019\\maven\\testdata\\lucky.xlsx";
	
	File f=new File(filepath);
	
	  Workbook wb=null;
	  Sheet sheet=null;
	  Row row=null;
	  Cell cell=null;
	  
	  if(filepath.endsWith(".xls"))
		  wb=new HSSFWorkbook();
	  else if(filepath.endsWith(".xlsx"))
		  wb=new XSSFWorkbook();
	  
	  sheet=wb.createSheet("test");
	  
	  for(int i=0;i<10;i++)
	  {		  row=sheet.createRow(i);
	  for(int j=0;j<10;j++)
	  {  cell=row.createCell(j);
	  cell.setCellValue("test"+i+j);
	  }
	  }  
	  FileOutputStream fos=new FileOutputStream(f);
	  wb.write(fos);
	  wb.close();
	  
  }
}
