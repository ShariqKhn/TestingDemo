package pack_1;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFile 
{
	public static void main(String[] args) 
	{
		try
		{
		FileInputStream fi=new FileInputStream(new File("C:\\Users\\pc\\eclipse-workspace\\Project_1\\excel\\data.xlsx"));
		XSSFWorkbook book=new XSSFWorkbook(fi);
		
		XSSFSheet sheet=book.getSheet("sheet1");
		
		String celldata=sheet.getRow(0).getCell(0).getStringCellValue();
		double data1=sheet.getRow(1).getCell(1).getNumericCellValue();
		System.out.println(celldata);
		System.out.println(data1);
		}
		catch(Exception e)
		{
			e.getMessage();
			e.getCause();
			e.printStackTrace();
		}
		
	
		
	/*	for(int i=0;i<r;i++)
		{
			XSSFRow xf=sheet.getRow(i);
			
			for(int j=0;j<xf.getPhysicalNumberOfCells();j++)
			{
				
			
				XSSFCell xc=xf.getCell(j);
				System.out.println(xc.getStringCellValue());
				
			}
}*/
		
	
		
	}

	

}
