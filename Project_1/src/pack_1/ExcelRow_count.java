package pack_1;

import java.io.File;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRow_count 
{
	public static void getrow_count()
	{
		try
		{
		/*String url=System.getProperty("user.dir");
		XSSFWorkbook book=new XSSFWorkbook(url+"/excel/data.xlsx");*/
		XSSFWorkbook book=new XSSFWorkbook(new File("C:\\Users\\pc\\eclipse-workspace\\Project_1\\excel\\data.xlsx"));
		XSSFSheet sheet=book.getSheet("Sheet1");
			
		int row_count=sheet.getPhysicalNumberOfRows();
		System.out.println("Number of Rows :- " +row_count);
			
		}
		catch(Exception e)
		{
		System.out.println(e.getMessage());
		System.out.println(e.getCause());
		e.printStackTrace();
			}
		}
	public static void main(String[] args) 
	{
		getrow_count();
		
	}

}
