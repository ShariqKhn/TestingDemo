package pack2;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteFile 
{
	public static void main(String[] args) 
	{
		try
		{
		String url=System.getProperty("user.dir");
		
		FileOutputStream fo=new FileOutputStream(new File(url+"/excel/file1.xlsx"));
		XSSFWorkbook book=new XSSFWorkbook();
		
	    XSSFSheet sheet=book.createSheet("Sheet1");

	    for(int i=0;i<4;i++)
	    {
	    	XSSFRow sr=sheet.createRow(i);
	    	
	    	for(int j=0;j<4;j++)
	    	{
	    		XSSFCell xc=sr.createCell(j);
	    		xc.setCellValue("Shariq");
	    		xc.setCellValue("khan"); // It will print only "khan"
	    	}
	    }
	    
	    book.write(fo);
	    fo.flush();
	    book.close();
	    fo.close();
}
		catch(Exception e)
		{
			e.getMessage();
			e.getCause();
			e.printStackTrace();
		}
		
		
		
	}

}
