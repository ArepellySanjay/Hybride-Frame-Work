package AppUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XL_Utils_Next 
{



	public static 	FileInputStream fi; 
	public static  Workbook wb;
	public static  Sheet ws;
	public static   Row row;
	public static  Cell cell;

	public static int getRowcount(String xlfile, String xlsheet) throws IOException 
	{
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws =wb.getSheet(xlsheet);
		int rowcount =    ws.getLastRowNum();	
		wb.close();
		return rowcount;

	}




	public static  short getcoloumncount(String xlfile, String xlsheet, int rownum) throws IOException 
	{
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws =	wb.getSheet(xlsheet);
		row =	ws.getRow(rownum);

		int colcount =	row.getLastCellNum();
		wb.close();
		return (short) colcount;

	}


	public  static  String  getStringData(String xlfile ,String xlsheet, int rowcont ,int celcount) throws IOException 		
	{

		fi= new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row= ws.getRow(rowcont);


		String data1 ;
		try {

			cell =	row.getCell(celcount);
			data1 = cell.getStringCellValue();

		} catch (Exception e) 
		{
			data1 = " ";
		}
 

		wb.close();
		return data1;
	}






	public static  double getNumaricData(String Xlfile, String xlsheet ,int rownum , int cellcount ) throws IOException 
	{

		fi = new FileInputStream(Xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);

		double salary;
		try {

			cell = row.getCell(cellcount);
			salary =	cell.getNumericCellValue();

		} catch (Exception e) 
		{
			salary = 0.0;
		}

		wb.close();
		return salary;

	}


	public static boolean getbooleanData(String xlfile ,String xlsheet , int rownum , int colcount) throws IOException
	{

		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);

		boolean status;
		try {

			cell  =	row.getCell(colcount);

			status =	cell.getBooleanCellValue();


		} catch (Exception e) {

			status = false;

		}
		wb.close();
		return status;
	}	


	public static void getsetData(String xlfile , String xlsheet, int rownum , int colcount , String data1) throws IOException 
	{


		fi =  new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row =ws.getRow(rownum);
		cell=  row.createCell(colcount);



		cell.setCellValue(data1);

		FileOutputStream fo = new FileOutputStream(xlfile);
		wb.write(fo);
		wb.close();


	}

	public static void FillGreenColour(String xlfile , String xlsheet , int rownum ,int colcount) throws IOException // string output
	{

		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		cell =row.getCell(colcount);

		//cell.setCellValue(pass);

		CellStyle style = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell.setCellStyle(style);

		FileOutputStream fo = new FileOutputStream(xlfile);
		wb.write(fo);
		wb.close();
	}
	
	public static void FillRedColour(String xlfile , String xlsheet , int rownum ,int colcount) throws IOException // string output
	{

		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		cell =row.getCell(colcount);

		// cell.setCellValue(pass);

		CellStyle style = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.RED.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell.setCellStyle(style);

		FileOutputStream fo = new FileOutputStream(xlfile);
		wb.write(fo);
		wb.close();



	}


	public static  String getStringSanjuData(String xlfile, String xlsheet, int rownum, int colcount) throws IOException
	{

		fi = new FileInputStream(xlfile);
		wb= new XSSFWorkbook(fi);
		ws= wb.getSheet(xlsheet);
		row = ws.getRow(rownum);


		String sanju;
		try {

			cell = row.getCell(colcount);
			sanju =	cell.getStringCellValue();

		} catch (Exception e) 
		{

			sanju = "";

		}

		return sanju;

	}




}





