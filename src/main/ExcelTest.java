package main;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

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
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import main.util.CreditInfoExcelUtil;

public class ExcelTest {

	public static void main(String[] args) throws IOException {
		
		CreditInfoExcelUtil excelUtil = new CreditInfoExcelUtil("xlsx");
		
		
		String[] usedHeader = { "이용내역","이용내역2","이용내역3","이용내역4","이용내역5"};
		String[][] usedContents = { { "내용입니다.", "내용",
			"Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.",
				"2022-01-01\n2022-01-02\n2022-01-03", "안내메세지 \n 입니당당당"}};
		
		excelUtil.setTable(CreditInfoExcelUtil.usedStr, usedHeader, usedContents, 2, 2);
		
		String filePath = "mail/";
		DateFormat dateFormat = new SimpleDateFormat("yyyyMMddHHmmss");
		String dateToStr = dateFormat.format(new Date());

		String fullExcelFileName = "MAIL_"+dateToStr;
		
		excelUtil.writeExcel(filePath, fullExcelFileName);
		excelUtil.closeWorkbook();
	}

	
}
