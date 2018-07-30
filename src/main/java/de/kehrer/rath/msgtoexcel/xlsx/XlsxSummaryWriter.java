package de.kehrer.rath.msgtoexcel.xlsx;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Calendar;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;
import org.springframework.util.Assert;

import de.kehrer.rath.msgtoexcel.msg.pojo.MsgContent;
import lombok.extern.slf4j.Slf4j;

@Slf4j
@Component
public class XlsxSummaryWriter {

	private static final String TEMPLATE_XLSX = "/template.xlsx";
	public static final String DEFAULT_EXCEL_DIR = "./excel";
	public static final String EXCEL_DIR_KEY = "EXCEL_DIR";

	private static final String[] COLUMNS = { "Nachname", "Vorname", "Straße-Haus-Nr.", "PLZ", "Ort", "Email", "Code",
			"Eingang" };

	@Value("${" + EXCEL_DIR_KEY + ":" + DEFAULT_EXCEL_DIR + "}")
	private Path excelFilesDir;

	public void writeSummary(String fileName, List<MsgContent> messages) {
		Assert.notNull(fileName, "fileName must not be null");
		Assert.notNull(messages, "msgContent must not be null");

		InputStream inputStream = getClass().getResourceAsStream(TEMPLATE_XLSX);
		
		try (Workbook workbook = new XSSFWorkbook(inputStream)) {
			CreationHelper createHelper = workbook.getCreationHelper();
			Sheet sheet = workbook.getSheet("Rezeptbücher");

			writeMessages(messages, workbook, createHelper, sheet);

			resize(sheet);
			writeOutputFile(fileName, workbook);
		} catch (Exception e) {
			log.error("Error writing the Excel file.", e);
		}
	}

	private void writeOutputFile(String fileName, Workbook workbook) throws FileNotFoundException, IOException {
		Files.createDirectories(excelFilesDir);
		FileOutputStream fileOut = new FileOutputStream(new File(excelFilesDir.toFile(), fileName));
		workbook.write(fileOut);
		fileOut.close();
	}

	private void resize(Sheet sheet) {
		for (int i = 0; i < COLUMNS.length; i++) {
			sheet.autoSizeColumn(i);
		}
	}

	private void writeMessages(List<MsgContent> messages, Workbook workbook, CreationHelper createHelper, Sheet sheet) {
		
		
		
		
		int rowNum = 3;
		for (MsgContent actContent : messages) {
			Row row = sheet.getRow(rowNum++);

			addCell(row, 0, actContent.getAddress().getLastName());
			addCell(row, 1, actContent.getAddress().getFirstName());
			addCell(row, 2, actContent.getAddress().getStreetAndHouseNumber());
			addCell(row, 3, Integer.parseInt(actContent.getAddress().getZipCode()));
			
			addCell(row, 4, actContent.getAddress().getCity());
			addCell(row, 5, actContent.getEMail());
			addCell(row, 6, actContent.getVoucherCode());
			addCell(row, 7, actContent.getMessageDate(),createHelper);
		}
	}
	
	private Cell addCell(Row row, int column, Integer value) {
		Cell createCell = row.getCell(column);
		createCell.setCellValue(value);		
		return createCell;
	}

	private Cell addCell(Row row, int column, Calendar value, CreationHelper createHelper) {
		Cell createCell = row.getCell(column);		
		short format = createHelper.createDataFormat().getFormat("dd.MM.yyyy");
		createCell.getCellStyle().setDataFormat(format);
		createCell.setCellValue(value);
		return createCell;
	}

	private Cell addCell(Row row, int column, String value) {
		Cell createCell = row.getCell(column);
		createCell.setCellValue(value);		
		return createCell;
	}
	
}
