package pack3;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadRowCount 
{
	public static void getRowCount(int rowNum,int column)
	{
		try
		{
		FileInputStream fi=new FileInputStream(new File("C:/Users/pc/eclipse-workspace/Project_1/excel/data.xlsx"));
		
		XSSFWorkbook book=new XSSFWorkbook(fi);
		XSSFSheet sheet=book.getSheetAt(0);
		
		int rowcount=sheet.getLastRowNum()-sheet.getFirstRowNum();
		
		 for (int i = 0; i < rowcount + 1; i++) {
		        Row row = sheet.getRow(i);

		        // create a loop to print cell values

		        for (int j = 0; j < row.getLastCellNum(); j++) {
		            Cell cell = row.getCell(j);
		           switch (cell.getCellType()) {
				//case Cell.CELL_TYPE_STRING:
					
					break;


		          //  case Cell.CELL_TYPE_NUMERIC:
		                System.out.print((int)row.getCell(j).getNumericCellValue() + " ");
		                break;

		            }

		        }

		        System.out.println();


		 }	
		
		 }
		
		
		/*String celldata=sheet.getRow(rowNum).getCell(column).getStringCellValue();
	    System.out.println(celldata);*/
		
        
		catch(Exception e)
		{
			e.getMessage();
			e.getCause();
			e.printStackTrace();
		}
}
	public static void main(String[] args) 
	{
		
		getRowCount(1,1 );
		
	}
	

}
