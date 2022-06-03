package main.util;

import java.io.FileOutputStream;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {
	
	public Workbook createWorkbook(String version) {
		if ("xls".equals(version)) {
			return new HSSFWorkbook();
		} else if ("xlsx".equals(version)) {
			return new XSSFWorkbook();
		}
		throw new NoClassDefFoundError();
	}
	public Row getRow(Sheet sheet, int rownum) {
		Row row = sheet.getRow(rownum);
		if (row == null) {
			row = sheet.createRow(rownum);
		}
		return row;
	}

	public Cell getCell(Row row, int cellnum) {
		Cell cell = row.getCell(cellnum);
		if (cell == null) {
			cell = row.createCell(cellnum);
		}
		return cell;
	}

	public Cell getCell(Sheet sheet, int rownum, int cellnum) {
		Row row = getRow(sheet, rownum);
		return getCell(row, cellnum);
	}

	public void writeExcel(Workbook workbook, String filepath) {
		try {
			FileOutputStream stream = new FileOutputStream(filepath);
			workbook.write(stream);
		} catch (Throwable e) {
			e.printStackTrace();
		}
	}
	public String strToDate(String strDate) {
		String result ="";
		if(strDate==null) {
			return result;
		} else {
			try {
		        SimpleDateFormat df;
		        SimpleDateFormat ndf;
		        if(strDate.length() == 14) {
		        	df = new SimpleDateFormat("yyyyMMddHHmmss");
		        	ndf = new SimpleDateFormat("yyyy/MM/dd HH:mm");
		        } else {// if(strDate.length() == 8) {
		        	df = new SimpleDateFormat("yyyyMMdd");
		        	ndf = new SimpleDateFormat("yyyy/MM/dd");
		        }
		        Date date = df.parse(strDate);
		        result = ndf.format(date);
		    } catch (ParseException e) {
		        return strDate;
		    }
		}
		return result;
	}
	public String strToAMT(Double douAMT, String curCode) {
		if(curCode.equals("KRW")) {
			DecimalFormat formatter = new DecimalFormat("###,###");
			return formatter.format(douAMT);
			
		} else if(curCode.equals("USD")) {
			DecimalFormat formatter = new DecimalFormat("###,###.####");
			return formatter.format(douAMT);
		}
		return douAMT.toString();
	}
	public String strToBirth(String strDate) {
		String result="";
		try {
	        SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd");
	        SimpleDateFormat ndf = new SimpleDateFormat("yyyy.MM.dd");
	        Date date = df.parse(strDate);
	        result = ndf.format(date);
	    } catch (ParseException e) {
	        return strDate;
	    }
		return result;
	}
	public CellStyle getCellStyle(Workbook workbook, String type) {
		CellStyle style = workbook.createCellStyle();
		
		if(type.equals("header")) {
			// Header Font
			style.setAlignment(HorizontalAlignment.CENTER);
			style.setFillForegroundColor(IndexedColors.GREY_80_PERCENT.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			
			Font white = workbook.createFont();
			white.setColor(IndexedColors.WHITE.getIndex());
			style.setFont(white);
		} else if(type.equals("title")) {
			Font bold14 = workbook.createFont();
			bold14.setFontHeight((short)(14*20)); 
			bold14.setBold(true);
			style.setFont(bold14);
		} else if(type.equals("date")) {
			// Date
			short df = workbook.createDataFormat().getFormat("MM/dd/yy HH:mm");
			style.setDataFormat(df);
		} else if(type.equals("krw")) {
			// KRW
			short df = workbook.createDataFormat().getFormat("#,##0");
			style.setDataFormat(df);
		} else if(type.equals("usd")) {
			// USD
			short df = workbook.createDataFormat().getFormat("#,##0.0000");
			style.setDataFormat(df);
		}
		
		return style; 
	}
	public void setTable(Workbook workbook, Sheet sheet, String[] header, String[][] contents, int cellC, int rowCount) {
		System.out.println("setTable");
	    Cell cell = null;
	    int cellCount = 0;
	    Row row = null;

	    // Header
	    row = sheet.createRow(rowCount++);
	    cellCount = cellC;
	    for(int i=0;i<header.length;i++) {
	        cell = row.createCell(cellCount++);
	        cell.setCellStyle(getCellStyle(workbook, "header"));
	        cell.setCellValue(header[i]);
	    }
	    
	    // List
	    for(int i=0;i<contents.length;i++) {
	        row = sheet.createRow(rowCount++);
	        cellCount = cellC;
	        
	        for(int j=0;j<contents[i].length;j++) {
	            cell = row.createCell(cellCount++);
	            cell.setCellValue(contents[i][j]);
	        }
	    }
	}
}
