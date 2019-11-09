package pack3;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CopyFile 
{
	public static void main(String[] args) 
	{
		try
		{
		File file=new File("C:/Users/pc/eclipse-workspace/Project_1/excel/data.xlsx");	
			
		FileInputStream fi=new FileInputStream(file);
		
		XSSFWorkbook book=new XSSFWorkbook(fi);
		
		XSSFSheet sheet=book.getSheetAt(0);
		
		//int r=sheet.getPhysicalNumberOfRows();
		FileOutputStream fo=new FileOutputStream(new File("C:/Users/pc/eclipse-workspace/Project_1/excel/demo.xlsx"));
		book.write(fo);
		fo.flush();
		fo.close();
		fi.close();
		book.close();
                System.out.println("Changes has done..!!")	
		System.out.println("Sucessfully copied");
		
}
		catch(Exception e)
		{
			e.getMessage();
			e.getCause();
			e.printStackTrace();
		}
	
	}











