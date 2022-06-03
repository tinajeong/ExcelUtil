/**
 * 
 */
package main.util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * @author USER
 *
 */
public class CreditInfoExcelUtil extends ExcelUtil {

	private Workbook workbook;
	private String version;
	
	private Sheet usedSheet;
	private Sheet providedSheet;

	public static final String usedStr = "이용내역";
	public static final String providedStr = "제공내역";

	private CellStyle headerStyle;
	private CellStyle headerBodyStyle;
	private CellStyle bodyStyle;

	public CreditInfoExcelUtil() {
		super();
	}

	public CreditInfoExcelUtil(String version) {
		this.initExcel(version);
	}

	public void initExcel(String version) {
		this.workbook = this.createWorkbook(version);
		this.version = version;
		this.initSheet();
		this.initStyle();
	}

	@Override
	public Workbook createWorkbook(String version) {
		this.workbook = super.createWorkbook(version);
		return this.workbook;
	}

	private void initSheet() {
		this.usedSheet = this.workbook.createSheet(usedStr);

		this.usedSheet.setColumnWidth(2, 8000);
		this.usedSheet.setColumnWidth(3, 4000);
		this.usedSheet.setColumnWidth(4, 10000);
		this.usedSheet.setColumnWidth(5, 6000);
		this.usedSheet.setColumnWidth(6, 8000);

		this.providedSheet = this.workbook.createSheet(providedStr);

		this.providedSheet.setColumnWidth(0, 8000);
		this.providedSheet.setColumnWidth(1, 4000);
		this.providedSheet.setColumnWidth(2, 10000);
		this.providedSheet.setColumnWidth(3, 6000);
		this.providedSheet.setColumnWidth(4, 8000);
	}

	private void initStyle() {

		this.headerStyle = workbook.createCellStyle();

		headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		headerStyle.setBorderRight(BorderStyle.MEDIUM);
		headerStyle.setBorderLeft(BorderStyle.MEDIUM);
		headerStyle.setBorderBottom(BorderStyle.MEDIUM);
		headerStyle.setBorderTop(BorderStyle.MEDIUM);

		Font font = workbook.createFont();
		font.setBold(true);

		headerStyle.setFont(font);

		this.headerBodyStyle = workbook.createCellStyle();

		headerBodyStyle.setWrapText(true);
		headerBodyStyle.setBorderRight(BorderStyle.THIN);
		headerBodyStyle.setBorderLeft(BorderStyle.THIN);
		headerBodyStyle.setBorderBottom(BorderStyle.DOUBLE);
		headerBodyStyle.setBorderTop(BorderStyle.THIN);
		headerBodyStyle.setAlignment(HorizontalAlignment.CENTER);
		headerBodyStyle.setVerticalAlignment(VerticalAlignment.CENTER);

		this.bodyStyle = workbook.createCellStyle();

		bodyStyle.setWrapText(true);
		bodyStyle.setBorderRight(BorderStyle.THIN);
		bodyStyle.setBorderLeft(BorderStyle.THIN);
		bodyStyle.setBorderBottom(BorderStyle.THIN);
		bodyStyle.setBorderTop(BorderStyle.THIN);
		bodyStyle.setAlignment(HorizontalAlignment.CENTER);
		bodyStyle.setVerticalAlignment(VerticalAlignment.CENTER);

	}

	public CellStyle getCellStyle(String type) {
		if (type.equals("header")) {
			return headerStyle;
		} else if (type.equals("header_body")) {
			return headerBodyStyle;
		} else { 
			return bodyStyle;
		}
	}

	public void setTable(String sheetName, String[] header, String[][] contents, int cellC, int rowCount) {

		System.out.println("setTable");

		Cell cell = null;
		int cellCount = 0;
		Row row = null;
		Sheet sheet = sheetName.equals(usedSheet.getSheetName()) ? this.usedSheet : this.providedSheet;

		// Title
		row = sheet.createRow(rowCount++);
		
		cellCount = cellC;
		for (int i = 0; i < header.length; i++) {
			cell = row.createCell(cellCount++);
			cell.setCellStyle(this.getCellStyle("header"));
			cell.setCellValue(sheetName);
		}
		sheet.addMergedRegion(new CellRangeAddress(cellC, cellC, cellC, cellC + header.length - 1));

		// Header
		row = sheet.createRow(rowCount++);
		cellCount = cellC;
		for (int i = 0; i < header.length; i++) {
			cell = row.createCell(cellCount++);
			cell.setCellStyle(this.getCellStyle("header_body"));
			cell.setCellValue(header[i]);
		}

		// List
		for (int i = 0; i < contents.length; i++) {
			row = sheet.createRow(rowCount++);
			cellCount = cellC;

			for (int j = 0; j < contents[i].length; j++) {
				cell = row.createCell(cellCount++);
				cell.setCellStyle(this.getCellStyle("body"));
				cell.setCellValue(contents[i][j]);
			}
		}
	}

	public String dateArrayToStr(String[] dateArr) {
		StringBuffer sb = new StringBuffer();

		for (int i = 0; i < dateArr.length; i++) {
			sb.append(this.strToDate(dateArr[i]) + "\n");
		}

		return sb.toString();
	}

	public void writeExcel(String filePath, String excelFileName) {
		System.out.println("WriteExcel : " + this.workbook + ", " + filePath);

		
		File file = new File(filePath);
		if (!file.isDirectory()) {
			file.mkdirs();
		}

		try {
			
			System.out.println(this.workbook.getSpreadsheetVersion());
			String fullExcelFileName = filePath + excelFileName + "." + this.version;
			FileOutputStream stream = new FileOutputStream(fullExcelFileName);
			
			workbook.write(stream);

		} catch (Throwable e) {
			e.printStackTrace();
		}

	}

	public void closeWorkbook() throws IOException {
		this.workbook.close();
	}

}
