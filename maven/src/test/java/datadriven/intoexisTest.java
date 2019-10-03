package datadriven;

import org.testng.annotations.Test;
import java.io.File;
import java.io.FileInputStream;
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

public class intoexisTest {
	 Workbook wb;
	 Sheet sh;
	 Row row;
	 Cell cell;
  @Test
  public void f() throws IOException {
	  
	  String filepath="E:\\seleniumworkspace2019\\maven\\testdata\\test.xlsx";
	 
	  String sheet="testcases";
	 
	File f=new File(filepath);

	 FileInputStream fip=new FileInputStream(f);
	
	 if(filepath.endsWith(".xls"))
	 {
		 wb =new HSSFWorkbook()	;
	 } 
	 else if(filepath.endsWith(".xlsx"))
	 { 
		 wb=new XSSFWorkbook();
	 }
	 
	 sh=wb.getSheet("testcases");
	 int nr=sh.getLastRowNum()+1;
	 for(int i=0; i<nr; i++)
	 {
		 row=sh.getRow(i);
		 cell=row.getCell(5);
		 if(cell==null)
		 {
			 cell=row.createCell(5);
			 cell.setCellValue("pass");
		 }
	 }
		 
		 FileOutputStream fos=new FileOutputStream(f);
		 
		  wb.write(fos);
		  
		  wb.close();
	 
			 
	 
	  
	  
	  
  }
}
