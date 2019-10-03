package datadriven;


import org.testng.annotations.Test;
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

	  public class intoexistingfileTest {
	  @Test
	    public void f() throws IOException {
	  	  String path="E:\\seleniumworkspace2019\\maven\\testdata\\test.xlsx";
	  	  String sheet="testcases";
	  	  File f=new File(path);
	  	  FileInputStream fis=new FileInputStream(f);
	  	  Workbook wb=null;
	  	  Sheet sh=null;
	  	  Row r=null;
	  	  Cell cell=null;
	  	  
	  	  if(path.endsWith(".xls"))
	  		  wb=new HSSFWorkbook(fis);
	  	  else if(path.endsWith(".xslx"))
	  		  wb=new XSSFWorkbook(fis);
	  	      sh = wb.getSheet("testcases");
	  	  int nr=sh.getLastRowNum()+1;
	  	  for(int i=1;i<nr;i++)
	  	  {
	  		  r=sh.getRow(i);
	  	  cell=r.getCell(5);
	  	  if(cell==null)
	  			cell=r.createCell(5);
	  		
	  		cell.setCellValue("PASS");
	  	  }	
	  		FileOutputStream fos = new FileOutputStream(f);
	  		wb.write(fos);
	  		wb.close();
	  		
	  	}
	  	
	  	  	  
  
}
