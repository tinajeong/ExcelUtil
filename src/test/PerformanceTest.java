/**
 * 
 */
package test;

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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class PerformanceTest {

	@Test(timeout=2000)
	public void setStyleTest() throws IOException {
		long startTime = System.currentTimeMillis();
		XSSFWorkbook workbook = new XSSFWorkbook();
		CellStyle style = workbook.createCellStyle();
		Sheet sheet = workbook.createSheet("test");
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);

		for (int i = 0; i < 10000; i++) {
			cell.setCellValue("test");
			style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			style.setBorderRight(BorderStyle.MEDIUM);
			style.setBorderLeft(BorderStyle.MEDIUM);
			style.setBorderBottom(BorderStyle.MEDIUM);
			style.setBorderTop(BorderStyle.MEDIUM);
			cell.setCellStyle(style);
		}

		workbook.close();

		System.out.println(System.currentTimeMillis() - startTime);
	}

	@Test(timeout=3000)
	public void baseTest() throws IOException {
		long startTime = System.currentTimeMillis();
		XSSFWorkbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("test");
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);

		for (int i = 0; i < 10000; i++) {
			cell.setCellValue("test");

			CellStyle style = workbook.createCellStyle();

			style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			style.setBorderRight(BorderStyle.MEDIUM);
			style.setBorderLeft(BorderStyle.MEDIUM);
			style.setBorderBottom(BorderStyle.MEDIUM);
			style.setBorderTop(BorderStyle.MEDIUM);

			cell.setCellStyle(style);
		}

		workbook.close();

		System.out.println(System.currentTimeMillis() - startTime);
	}

	@Test(timeout=2000)
	public void predefinedStyleTest() throws IOException {
		long startTime = System.currentTimeMillis();

		XSSFWorkbook workbook = new XSSFWorkbook();

		Sheet sheet = workbook.createSheet("test");
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);

		CellStyle header = workbook.createCellStyle();
		header.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		header.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		header.setBorderRight(BorderStyle.MEDIUM);
		header.setBorderLeft(BorderStyle.MEDIUM);
		header.setBorderBottom(BorderStyle.MEDIUM);
		header.setBorderTop(BorderStyle.MEDIUM);

		Font font = workbook.createFont();
		font.setBold(true);
		header.setFont(font);

		CellStyle headerBody = workbook.createCellStyle();
		headerBody.setWrapText(true);
		headerBody.setBorderRight(BorderStyle.THIN);
		headerBody.setBorderLeft(BorderStyle.THIN);
		headerBody.setBorderBottom(BorderStyle.DOUBLE);
		headerBody.setBorderTop(BorderStyle.THIN);
		headerBody.setAlignment(HorizontalAlignment.CENTER);
		headerBody.setVerticalAlignment(VerticalAlignment.CENTER);

		CellStyle body = workbook.createCellStyle();
		body.setWrapText(true);
		body.setBorderRight(BorderStyle.THIN);
		body.setBorderLeft(BorderStyle.THIN);
		body.setBorderBottom(BorderStyle.THIN);
		body.setBorderTop(BorderStyle.THIN);
		body.setAlignment(HorizontalAlignment.CENTER);
		body.setVerticalAlignment(VerticalAlignment.CENTER);

		for (int i = 0; i < 10000; i++) {
			cell.setCellValue("test");
			if (i % 3 == 0) {
				cell.setCellStyle(header);
			} else if (i % 3 == 1) {
				cell.setCellStyle(headerBody);
			} else {
				cell.setCellStyle(body);
			}
		}

		workbook.close();

		System.out.println(System.currentTimeMillis() - startTime);
	}
}
