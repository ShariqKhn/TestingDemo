package pack2;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteData 
{
	public synchronized void function()  
	{
		try
		{
	
		File file=new File("C:\\Users\\pc\\eclipse-workspace\\Project_1\\excel\\data_1.xlsx");
		FileInputStream fi=new FileInputStream(file);
		
			
			
			
			
		XSSFWorkbook book=new XSSFWorkbook();
		XSSFSheet sheet=book.getSheetAt(0);
		
		sheet.getRow(0).createCell(2).setCellValue("Values");
		sheet.getRow(1).createCell(2).setCellValue("pass");
		sheet.getRow(2).createCell(2).setCellValue("Fail");
		sheet.getRow(2).createCell(2).setCellValue("Avg");
		
		
		FileOutputStream fo=new FileOutputStream(file);
		
		book.write(fo);
		book.close();
		
	
		System.out.println("Sucessfull added..!!!");
}
		catch(Exception e)
		{
			e.getMessage();
			e.getCause();
			e.printStackTrace();
		}
	}
	public static void main(String[] args) 
	{
		WriteData wd=new WriteData();
		wd.function();
		
	}

}
