package excelcollect;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;

public class Main {
	/**
	 * Later I should read the directory path from a property file or something.
	 */
	private static final String DATA_DIR = "E:\\WorkSpaces\\data";
	private static final String OUTPUT_FILE = "E:\\WorkSpaces\\test.xlsx";
	private static final DataFormatter formatter = new DataFormatter();

	public static void main(String[] args) {
		System.out.println("Start processing");

		List<File> inputFiles = getFileList(DATA_DIR);
		String[] names = { "テーブル1", "テーブル2" };
		List<String> sheetNames = new ArrayList<String>();
		for (String name : names) {
			sheetNames.add(name);
		}

		XSSFWorkbook outputWb = null;
		try {
			FileInputStream fis = new FileInputStream(new File(OUTPUT_FILE));
			outputWb = new XSSFWorkbook(fis);
			XSSFSheet sheet = outputWb.getSheet("Summary");
			XSSFTable table = sheet.getTables().get(0);

			for (File file : inputFiles) {
				appendToTable(table, readFile(file, sheetNames));
			}

			fis.close();
			FileOutputStream fos = new FileOutputStream(new File(OUTPUT_FILE));
			outputWb.write(fos);
			fos.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	/**
	 * Return a list of files in a given path. If the path does not exist, raise
	 * FileNotFoundException.
	 * 
	 * @param filePath
	 *            file path from which the list is taken
	 * @return list of files in the directory. Directories are not included.
	 */
	public static List<File> getFileList(String filePath) {
		List<File> fileList = new ArrayList<File>();

		File inputDir = new File(filePath);

		for (File file : inputDir.listFiles()) {
			if (file.isFile() && file.getName().endsWith(".xlsx")) {
				fileList.add(file);
			}
		}

		return fileList;
	}

	/**
	 * Return a list of CellValues within specified tables TODO: Decide what to
	 * do on error (such as the file is not an excel file, or the file does not
	 * exist)
	 * 
	 * @param inputFile
	 *            an excel File where the data is stored. It must be an .exlx
	 *            file so that this method can read data from a table.
	 * @param inputTableNames
	 *            a list of Strings of the tables where the data is stored
	 * @return
	 */
	public static List<Map<String, CellValue>> readFile(File inputFile, List<String> inputTableNames) {
		List<Map<String, CellValue>> data = new ArrayList<Map<String, CellValue>>();
		XSSFWorkbook wb = null;

		try {
			wb = new XSSFWorkbook(inputFile);
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		Iterator<Sheet> it = wb.sheetIterator();
		while (it.hasNext()) {
			XSSFSheet st = (XSSFSheet) it.next();
			List<XSSFTable> tables = st.getTables();

			if (tables == null || tables.size() == 0) {
				continue;
			}

			for (XSSFTable table : tables) {
				if (inputTableNames.contains(table.getName())) {
					data.addAll(readTable(table));
				}
			}
		}

		System.out.println("Read " + inputFile.getName());
		debugPrint(data);
		return data;
	}

	public static void debugPrint(List<Map<String, CellValue>> values) {
		for (Map<String, CellValue> map : values) {
			Set<String> keys = map.keySet();
			for (String key : keys) {
				System.out.println(key + "\t" + map.get(key).getValue());
			}
		}
	}

	public static List<Map<String, CellValue>> readTable(XSSFTable table) {
		List<Map<String, CellValue>> values = new ArrayList<Map<String, CellValue>>();
		CellReference startRef = table.getStartCellReference();
		CellReference endRef = table.getEndCellReference();

		List<String> headers = retrieveTableHeaders(table, startRef, endRef);

		for (int rowNum = startRef.getRow() + 1; rowNum <= endRef.getRow(); rowNum++) {
			Row row = table.getXSSFSheet().getRow(rowNum);

			Map<String, CellValue> rowData = new HashMap<String, CellValue>();
			// I don't know if `row.iterator()` returns cells in order of the
			// actual data,
			// so use column indices to retrieve values in correct order.
			for (short colNum = startRef.getCol(); colNum <= endRef.getCol(); colNum++) {
				int headerIndex = colNum - startRef.getCol();
				Cell cell = row.getCell(colNum);
				rowData.put(headers.get(headerIndex),
						new CellValue(cell.getCellType(), formatter.formatCellValue(cell)));
			}

			values.add(rowData);
		}

		return values;
	}

	/**
	 * Retrieve headers of a given table. The headers are sorted from left to
	 * right
	 * 
	 * @param table
	 *            an XSSF Table.
	 * @param startRef
	 *            a reference to the left top cell of the table
	 * @param endRef
	 *            a reference to the right-bottom cell of the table
	 * @return a list of headers of the table
	 */
	public static List<String> retrieveTableHeaders(XSSFTable table, CellReference startRef, CellReference endRef) {
		List<String> headers = new ArrayList<String>();

		Row row = table.getXSSFSheet().getRow(startRef.getRow());
		for (short colNum = startRef.getCol(); colNum <= endRef.getCol(); colNum++) {
			Cell cell = row.getCell(colNum);
			headers.add(cell.toString());
		}

		return headers;
	}

	/**
	 * Write values to a file. If there
	 * 
	 * @param filePath
	 * @param values
	 */
	public static void writeSummary(String filePath, List<Map<String, CellValue>> values) {
		XSSFWorkbook wb = null;
		try {
			wb = new XSSFWorkbook(new File(filePath));
			XSSFSheet st = wb.createSheet("Summary");
			XSSFTable table = st.createTable();

			table.setDisplayName("Summary");
			table.setName("Summary");
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * Append list to an existing table
	 * 
	 * @param table
	 * @param values
	 */
	public static void appendToTable(XSSFTable table, List<Map<String, CellValue>> values) {
		XSSFSheet sheet = table.getXSSFSheet();
		CellReference startRef = table.getStartCellReference();
		CellReference endRef = table.getEndCellReference();

		List<String> headers = retrieveTableHeaders(table, startRef, endRef);

		// append data to the sheet
		int appendingRowNum = endRef.getRow();
		for (Map<String, CellValue> map : values) {
			System.out.println("Writing keys: " + map.keySet().toString());
			Row row = sheet.getRow(++appendingRowNum);
			if (row == null) {
				row = sheet.createRow(appendingRowNum);
			}
			for (short colNum = startRef.getCol(); colNum <= endRef.getCol(); colNum++) {
				int headerIndex = colNum - startRef.getCol();
				Cell cell = row.createCell(colNum);
				updateCellVal(cell, map.get(headers.get(headerIndex)));
			}
		}

		// expand the table
		CTTable cttable = table.getCTTable();
		cttable.setRef(calcRef(startRef, appendingRowNum, endRef.getCol()));

		updateCellRefs(table);
	}

	/**
	 * Do very tricky things to update CellReferences of the table object. After
	 * adding some rows to the table, I need to update an underlying CTTable
	 * instance to change the area of the table. However, it won't update
	 * CellReference instances of the table.
	 * 
	 * @param table
	 * @throws SecurityException
	 * @throws NoSuchFieldException
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 */
	public static void updateCellRefs(XSSFTable table) {
		try {
			Field startF = table.getClass().getDeclaredField("startCellReference");
			Field endF = table.getClass().getDeclaredField("endCellReference");
			startF.setAccessible(true);
			endF.setAccessible(true);
			startF.set(table, null);
			endF.set(table, null);
			startF.setAccessible(false);
			endF.setAccessible(false);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void updateCellVal(Cell cell, CellValue cellVal) {
		if (cellVal == null) {
			cell.setCellValue("");
		} else {
			switch (cellVal.getType()) {
			case Cell.CELL_TYPE_NUMERIC:
				try {
					cell.setCellValue(Double.parseDouble(cellVal.getValue()));
				} catch (NumberFormatException e) {
					cell.setCellValue(cellVal.getValue());
				}
				break;
			case Cell.CELL_TYPE_BLANK:
				cell.setCellValue("");
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				cell.setCellValue(Boolean.parseBoolean(cellVal.getValue()));
				break;
			case Cell.CELL_TYPE_FORMULA:
				// TODO: should change the method based on the value type,
				// sometimes
				// it is numeric, and otherwise String.
				cell.setCellValue(cellVal.getValue());
				break;
			default:
				cell.setCellValue(cellVal.getValue());
			}
		}
	}

	public static String calcRef(CellReference startRef, int endRowNum, int endColNum) {
		StringBuilder sb = new StringBuilder();

		sb.append(startRef.formatAsString());
		sb.append(":");

		CellReference newEndRef = new CellReference(endRowNum, endColNum);
		sb.append(newEndRef.formatAsString());

		return sb.toString();
	}
}
