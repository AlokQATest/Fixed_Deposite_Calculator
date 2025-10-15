package Data_Driven_Testing_FD_Calculator;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelUtills {
	
	public static FileInputStream fi;
	public static FileOutputStream fo;
	public static XSSFWorkbook workbook;
	public static XSSFSheet sheet;
	public static XSSFRow row;
	public static XSSFCell cell;
	public static CellStyle style;
	
	public static int getRowCount(String xlFile, String xlSheet) throws IOException {
		fi = new FileInputStream(xlFile);
		workbook = new XSSFWorkbook(fi);
		sheet = workbook.getSheet(xlSheet);
		int row_Count = sheet.getLastRowNum();
		workbook.close();
		fi.close();
		return row_Count;
	}
	
	public static int getCellCount(String xlFile, String xlSheet) throws IOException {
		fi = new FileInputStream(xlFile);
		workbook = new XSSFWorkbook(fi);
		sheet = workbook.getSheet(xlSheet);
		int cell_Count = sheet.getRow(0).getLastCellNum();
		workbook.close();
		fi.close();
		return cell_Count;
		
		
	}
	public static String getCellData(String xFile, String xlSheet, int row_Num, int column_Num) throws IOException {
		fi = new FileInputStream(xFile);
		workbook = new XSSFWorkbook(fi);
		sheet = workbook.getSheet(xlSheet);
		row = sheet.getRow(row_Num);
		cell = row.getCell(column_Num);
		String data;
		
		try {
			DataFormatter formatter = new DataFormatter();
			data = formatter.formatCellValue(cell);
		}catch(Exception e) {
			data ="";
		}
		return data;
	}
	public static void setCellData(String xFile, String xlSheet, int row_Num, int column, String data) throws IOException {
		fi = new FileInputStream(xFile);
		workbook = new XSSFWorkbook(fi);
		sheet = workbook.getSheet(xlSheet);
		row = sheet.getRow(row_Num);
		
		cell = row.createCell(column);
		cell.setCellValue(data);
		fo = new FileOutputStream(xFile);
		workbook.write(fo);
		workbook.close();
		fi.close();
		fo.close();	
	}
	
	public static void setGreenColor(String xFile, String xlSheet, int row_Num, int column) throws IOException {
		fi = new FileInputStream(xFile);
		workbook = new XSSFWorkbook(fi);
		sheet = workbook.getSheet(xlSheet);
		row = sheet.getRow(row_Num);
		cell = row.getCell(column);
		
		style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		cell.setCellStyle(style);
		fo = new FileOutputStream(xFile);
		workbook.write(fo);
		workbook.close();
		fi.close();
		fo.close();
		
	}
	
	public static void setRedColor(String xFile, String xlSheet, int row_Num, int column) throws IOException {
		fi = new FileInputStream(xFile);
		workbook = new XSSFWorkbook(fi);
		sheet = workbook.getSheet(xlSheet);
		row = sheet.getRow(row_Num);
		cell = row.getCell(column);
		
		style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.RED.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		cell.setCellStyle(style);
		fo = new FileOutputStream(xFile);
		workbook.write(fo);
		workbook.close();
		fi.close();
		fo.close();
		
		
	}
	
	
}
