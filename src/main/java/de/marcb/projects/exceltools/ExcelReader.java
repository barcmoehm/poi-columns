package de.marcb.projects.exceltools;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Optional;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelReader {

	private final String path;
	private final Workbook workbook;
	private final Sheet sheet;

	/**
	 * Creates a new Instance by trying to open a new workbook at the specified path. If no file is found, an
	 * IOException is thrown.
	 * Then the active sheet is opened.
	 *
	 * @param path
	 * @throws IOException
	 */
	public ExcelReader(final String path) throws IOException {
		this.path = path;
		this.workbook = open();
		this.sheet = workbook.getSheetAt(workbook.getActiveSheetIndex());
	}

	/**
	 * Creates a new Instance by trying to open a new workbook at the specified path. If no file is found, an
	 * IOException is thrown.
	 * Then the sheet with <i>sheetName</i> (or the active sheet, if the wanted sheet does not exist) is opened.
	 *
	 * @param path
	 * @throws IOException
	 */
	public ExcelReader(final String path, final String sheetName) throws IOException {
		this.path = path;
		this.workbook = open();
		this.sheet = openSheet(sheetName);
	}

	private Sheet openSheet(final String sheetName) {
		Sheet fallback = workbook.getSheetAt(workbook.getActiveSheetIndex());
		return Optional.ofNullable(workbook.getSheet(sheetName)).orElse(fallback);

	}

	/**
	 * Change active sheet to <i>sheetName</i>, returns false if the desired sheet does not exist.
	 */
	public boolean changeSheet(final String sheetName) {
		Sheet sheet = workbook.getSheet(sheetName);
		return sheet != null;
	}

	private Workbook open() throws IOException {
		File file = new File(path);
		FileInputStream inputStream = new FileInputStream(file);
		return new HSSFWorkbook(inputStream);
	}

	/**
	 * @param toFilter
	 *            Sheet (in column-representation) to be filtered
	 * @param filterRow
	 *            Name of the row that shall be filtered
	 * @param keyword
	 *            to filter for. Currently, the only available filter method is eq.
	 * @return A new Map with the filtered data
	 */
	// TODO accept T as filter citeria
	public Map<String, Column> filter(final Map<String, Column> toFilter, final String filterRow, final String keyword) {
		if (toFilter.get(filterRow) == null) {
			throw new NullPointerException("The desired row does not exist!");
		}
		List<Integer> foundRowIndizes = new ArrayList<>();
		Column searchForRowIndizes = toFilter.get(filterRow);
		for (int i = 0; i < searchForRowIndizes.count(); i++) {
			Cell cell = searchForRowIndizes.get(i);
			if (cell.getCellTypeEnum() == CellType.STRING && keyword.equalsIgnoreCase(cell.getStringCellValue())) {
				foundRowIndizes.add(i);
			}
		}
		Map<String, Column> filtered = new HashMap<>();
		for (Entry<String, Column> entry : toFilter.entrySet()) {
			filtered.put(entry.getKey(), new Column());
			for (int i = 0; i < entry.getValue().count(); i++) {
				if (foundRowIndizes.contains(i)) {
					filtered.get(entry.getKey()).add(entry.getValue().get(i));
				}
			}
		}

		return filtered;
	}

	/**
	 * Returns a column representation of the active sheet in the form of a map with Key(String) = column header and a
	 * list of the child columns
	 *
	 * @param headerRowIndex
	 * @return
	 */
	public Map<String, Column> columnRepresentation(final int headerRowIndex) {
		Map<Integer, Column> cols = createColumns(headerRowIndex + 1);
		Map<Integer, String> header = headerFromMap(headerRowIndex);
		return merge(header, cols);

	}

	private Map<Integer, Column> createColumns(final int startReadingAt) {
		final Map<Integer, Column> columns = new HashMap<>();
		int columnNumber = 0;

		for (int rowNumber = startReadingAt; rowNumber < sheet.getLastRowNum(); rowNumber++) {
			Row row = sheet.getRow(rowNumber);
			columnNumber = 0;
			for (int columnCount = 0; columnCount < row.getLastCellNum(); columnCount++) {
				if (!columns.containsKey(columnNumber)) {
					columns.put(columnNumber, new Column());
				}
				Cell cell = row.getCell(columnCount);
				if (cell != null) {
					columns.get(columnNumber).addIfAbsent(cell);
				}
				columnNumber++;
			}
		}

		return columns;
	}

	private Map<Integer, String> headerFromMap(final int headerRowIndex) {
		Row row = sheet.getRow(headerRowIndex);
		if (row == null) {
			throw new IllegalArgumentException("Required column not found!");
		}
		Map<Integer, String> header = new HashMap<>();
		int cellCount = 0;
		for (Cell cell : row) {
			header.put(cellCount, cell.getStringCellValue());
			cellCount++;
		}

		return header;

	}

	private Map<String, Column> merge(final Map<Integer, String> header, final Map<Integer, Column> columns) {
		Map<String, Column> merge = new HashMap<>();

		for (Entry<Integer, Column> column : columns.entrySet()) {
			merge.put(header.get(column.getKey()), column.getValue());
		}

		return merge;

	}

	public String getPath() {
		return path;
	}

	public Workbook getWorkbook() {
		return workbook;
	}

	public Sheet getSheet() {
		return sheet;
	}

}
