package pack_1;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excelutils 
{
	protected static String path;
	protected static XSSFWorkbook book;
	protected static XSSFSheet sheet;
	
	public Excelutils(String excelPath,String sheetName)
	{
	  try
		{
		    book=new XSSFWorkbook(excelPath);
			sheet=book.getSheet(sheetName);
			}
		catch(Exception e)
		{
			e.getMessage();
			e.getCause();
			e.printStackTrace();
		}
	}
	public static void getRowCount()
	{
		try
		{
		int rowcount=sheet.getPhysicalNumberOfRows();
		System.out.println("Number of Row in Sheet :- "+rowcount);
		}
		catch(Exception e)
		{
			e.getMessage();
			e.getCause();
			e.printStackTrace();
		}
	}
	
	public static void getCellDataString(int rowNum,int column)
	{
		try
		{
		String cellData=sheet.getRow(rowNum).getCell(column).getStringCellValue();
		System.out.println(cellData);
		}
		catch(Exception e)
		{
			e.getMessage();
			e.getCause();
			e.printStackTrace();
		}
}
	public static void getCellDataNumber(int rowNum,int column)
	{
		try
		{
		double cellNum=sheet.getRow(rowNum).getCell(column).getNumericCellValue();
		System.out.println(cellNum);
		}
		catch(Exception e)
		{
			e.getMessage();
			e.getCause();
			e.printStackTrace();
		}
	}
	
}
