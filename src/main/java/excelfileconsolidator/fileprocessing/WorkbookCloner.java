package excelfileconsolidator.fileprocessing;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class WorkbookCloner {

	private Workbook masterWorkbook; 

	public WorkbookCloner(Workbook workbook) {

		this.masterWorkbook = workbook;
	}

	public void clone(Workbook workbook, String fileName) {
		sheetProcessor(workbook, fileName);
	}

	private void sheetProcessor(Workbook workbook, String fileName) {

		int numSheets = workbook.getNumberOfSheets();
		for (int i = 0; i < numSheets; i++) {
			Sheet currentSheet = workbook.getSheetAt(i);
			Sheet masterSheet = masterWorkbook.createSheet(fileName + " - " + currentSheet.getSheetName());
			sheetReader(currentSheet, masterSheet);
		}
	}

	private void sheetReader(Sheet origSheet, Sheet destSheet) {
		int lastRow = origSheet.getLastRowNum();

		for (int j = 0; j <= lastRow; j++) {
			Row currentRow = origSheet.getRow(j);
			Row destRow = destSheet.createRow(j);
			rowReader(currentRow, destRow);
		}
	}

	private void rowReader(Row origRow, Row destRow) {

		Iterator<Cell> cells = origRow.cellIterator();
		while (cells.hasNext()) {
			Cell cell = cells.next();
			Cell destCell = null;
			switch (cell.getCellType()) {

			case BLANK:
				destCell = destRow.createCell(cell.getColumnIndex(), CellType.BLANK);
				destCell.setBlank();
				break;
			case BOOLEAN:
				destCell = destRow.createCell(cell.getColumnIndex(), CellType.BOOLEAN);
				destCell.setCellValue(cell.getBooleanCellValue());
				break;
			case FORMULA:
				destCell = destRow.createCell(cell.getColumnIndex(), CellType.FORMULA);
				destCell.setCellFormula(cell.getCellFormula());
				break;
			case NUMERIC:
				destCell = destRow.createCell(cell.getColumnIndex(), CellType.NUMERIC);
				destCell.setCellValue(cell.getNumericCellValue());
				break;
			case STRING:
				destCell = destRow.createCell(cell.getColumnIndex(), CellType.STRING);
				destCell.setCellValue(cell.getStringCellValue());
				break;
			default:
				System.out.println("Valor no reconocido");
			}
		}
	}

	public void write(String destPath) {

		String currentDate = new SimpleDateFormat("MM-DD-YYYY-HH-mm-ss").format(new Date());
		String fileSeparator = File.separator;
		File outputFile = new File(destPath + fileSeparator + "Master - " + currentDate + ".xlsx");
		try {

			FileOutputStream fileOutputStream = new FileOutputStream(outputFile);
			masterWorkbook.write(fileOutputStream);

			fileOutputStream.close();
			masterWorkbook.close();

		} catch (FileNotFoundException fnfex) {
			fnfex.printStackTrace();
		} catch (IOException ioex) {
			ioex.printStackTrace();
		}
	}
}