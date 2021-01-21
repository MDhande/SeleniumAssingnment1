package Cases;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils
{ 
	static String projectpath;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;

	public  ExcelUtils(String excelpath, String sheetname)
	{
		try {
			projectpath= System.getProperty("user.dir");
			workbook = new XSSFWorkbook(excelpath);
			sheet = workbook.getSheet(sheetname);	
		}
		catch(Exception e)
		{
			e.printStackTrace();	
		}
	}


	public static int getRow() {

		int rowcount=0;
		try {

			rowcount=sheet.getPhysicalNumberOfRows();
			System.out.println("no of rows are :" +rowcount);

		}
		catch(Exception e)
		{
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}
		return rowcount;

	}


	public static int getCol() {

		int Colcount=0;
		try {

			Colcount=sheet.getRow(0).getPhysicalNumberOfCells();
			System.out.println("no of columns are :" +Colcount);

		}
		catch(Exception e)
		{
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}
		return Colcount;
	}
	public static String getcellDataString(int rowNum, int colNum)
	{
		String cellData=null;
		try {


			cellData =sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
			//System.out.println(cellData);
		}

		catch(Exception ex)
		{
			System.out.println(ex.getMessage());
			System.out.println(ex.getCause());
			ex.printStackTrace();
		}
		return cellData;
	}

	public static void getcellDataNumber(int rowNum, int colNum)
	{
		try {

			double celldata =sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
			System.out.println(celldata);
		}

		catch(Exception ex)
		{
			System.out.println(ex.getMessage());
			System.out.println(ex.getCause());
			ex.printStackTrace();
		}
	}





}
