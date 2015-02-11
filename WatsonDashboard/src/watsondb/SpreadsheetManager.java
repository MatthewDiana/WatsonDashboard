package watsondb;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.NumberFormat;
import java.util.Collection;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SpreadsheetManager {

	private File inputFile;
	private File outputFile;
	private Dashboard db;
	
	public SpreadsheetManager(File inputFile, File outputFile) {
		
		this.inputFile = inputFile;
		this.outputFile = outputFile;
		db = new Dashboard();
		
	}
	
	public void read() {
		
		// Attempt to open input file.
		FileInputStream file = null;
		try {
			file = new FileInputStream(inputFile);
		} catch (FileNotFoundException e) {
			System.out.println("Unable to open " + inputFile.toString());
			e.printStackTrace();
		}
		
		// Attempt to open input workbook.
		XSSFWorkbook inputWorkbook = null;
		try {
			inputWorkbook = new XSSFWorkbook(file);
		} catch (IOException e) {
			System.out.println("Unable to open workbook from .xlsx file");
			e.printStackTrace();
		}
		
		// Create variables for parsing through sheets.
		Iterator<Row> rowIterator;
		Row headerRow;
		Person currentPerson;
		Department currentDepartment;
		
		// ===== PROPOSALS ===== //
		XSSFSheet proposals = inputWorkbook.getSheetAt(0);
		
		rowIterator = proposals.iterator();
		headerRow = rowIterator.next();
		if (headerRow == null) {
			System.out.println("Unable to locate header row.");
			System.exit(0);
		}
		
		currentPerson = null;
		currentDepartment = null;

		
		while (rowIterator.hasNext()) {
			int columnNum = 0;
			Cell headerCell = headerRow.getCell(columnNum);
			Row currentRow = rowIterator.next();
			Iterator<Cell> cellIterator = currentRow.iterator();
			while (cellIterator.hasNext()) {
				Cell currentCell = cellIterator.next();
				String currentCellStr = currentCell.toString();
				switch (headerCell.toString()) {
				case "Dept":
					String deptName = currentCellStr.substring(currentCellStr.indexOf(" ") + 1, currentCellStr.length());
					if (!db.hasDepartment(deptName))
						db.addNewDepartment(deptName);
					currentDepartment = db.getDepartment(deptName);
					break;
				case "Person":
					String personName = currentCellStr;
					if (!db.hasPerson(personName))
						db.addNewPerson(personName, currentDepartment);
					currentPerson = db.getPerson(personName);
					break;
				case "Direct Costs Credited":
					BigDecimal dcamount = new BigDecimal(currentCellStr);
					if (currentDepartment != null)
						currentDepartment.increaseProposalsTotal(dcamount);
					if (currentPerson != null)
						currentPerson.increaseProposalsTotal(dcamount);
					break;
				case "F&A Costs Credited":
					BigDecimal facamount = new BigDecimal(currentCellStr);
					if (currentDepartment != null)
						currentDepartment.increaseProposalsTotal(facamount);
					if (currentPerson != null)
						currentPerson.increaseProposalsTotal(facamount);
					break;
				}
				headerCell = headerRow.getCell(++columnNum);
			}
			currentPerson = null;
			currentDepartment = null;
		}
		
		// ===== EXPENDITURES ===== //
		XSSFSheet expenditures = inputWorkbook.getSheetAt(1);
		
		rowIterator = expenditures.iterator();
		headerRow = rowIterator.next();
		if (headerRow == null) {
			System.out.println("Unable to locate header row.");
			System.exit(0);
		}
		
		currentPerson = null;
		currentDepartment = null;

		
		while (rowIterator.hasNext()) {
			int columnNum = 0;
			Cell headerCell = headerRow.getCell(columnNum);
			Row currentRow = rowIterator.next();
			Iterator<Cell> cellIterator = currentRow.iterator();
			while (cellIterator.hasNext()) {
				Cell currentCell = cellIterator.next();
				String currentCellStr = currentCell.toString();
				switch (headerCell.toString()) {
				case "Credit Org Name":
					String deptName = currentCellStr.substring(currentCellStr.indexOf(" ") + 1, currentCellStr.length());
					if (!db.hasDepartment(deptName))
						db.addNewDepartment(deptName);
					currentDepartment = db.getDepartment(deptName);
					break;
				case "Person Full Name":
					String personName = currentCellStr;
					if (!db.hasPerson(personName))
						db.addNewPerson(personName, currentDepartment);
					currentPerson = db.getPerson(personName);
					break;
				case "Credit Total Expenditure":
					BigDecimal amount = new BigDecimal(currentCellStr);
					if (currentDepartment != null)
						currentDepartment.increaseExpendituresTotal(amount);
					if (currentPerson != null)
						currentPerson.increaseExpendituresTotal(amount);
				}
				headerCell = headerRow.getCell(++columnNum);
			}
			currentPerson = null;
			currentDepartment = null;
		}
		
		// ===== COMMITTED FUNDS ===== //
		XSSFSheet committedFunds = inputWorkbook.getSheetAt(2);
				
		rowIterator = committedFunds.iterator();
		headerRow = rowIterator.next();
		if (headerRow == null) {
			System.out.println("Unable to locate header row.");
			System.exit(0);
		}
		
		currentPerson = null;
		currentDepartment = null;

		
		while (rowIterator.hasNext()) {
			int columnNum = 0;
			Cell headerCell = headerRow.getCell(columnNum);
			Row currentRow = rowIterator.next();
			Iterator<Cell> cellIterator = currentRow.iterator();
			while (cellIterator.hasNext()) {
				Cell currentCell = cellIterator.next();
				String currentCellStr = currentCell.toString();
				switch (headerCell.toString()) {
				case "Award Organization Name":
					String deptName = currentCellStr.substring(currentCellStr.indexOf(" ") + 1, currentCellStr.length());
					if (!db.hasDepartment(deptName))
						db.addNewDepartment(deptName);
					currentDepartment = db.getDepartment(deptName);
					break;
				case "Award Principal Investigator":
					String personName = currentCellStr;
					if (!db.hasPerson(personName))
						db.addNewPerson(personName, currentDepartment);
					currentPerson = db.getPerson(personName);
					break;
				case "Fye Active Installment":
					BigDecimal amount = new BigDecimal(currentCellStr);
					if (currentDepartment != null)
						currentDepartment.increaseCommittedFundsTotal(amount);
					if (currentPerson != null)
						currentPerson.increaseCommittedFundsTotal(amount);
				}
				headerCell = headerRow.getCell(++columnNum);
			}
			currentPerson = null;
			currentDepartment = null;
		}
		
		// ===== NEW AWARDS ===== //
		XSSFSheet newAwards = inputWorkbook.getSheetAt(3);
				
		rowIterator = newAwards.iterator();
		headerRow = rowIterator.next();
		if (headerRow == null) {
			System.out.println("Unable to locate header row.");
			System.exit(0);
		}
		
		currentPerson = null;
		currentDepartment = null;

		
		while (rowIterator.hasNext()) {
			int columnNum = 0;
			Cell headerCell = headerRow.getCell(columnNum);
			Row currentRow = rowIterator.next();
			Iterator<Cell> cellIterator = currentRow.iterator();
			while (cellIterator.hasNext()) {
				Cell currentCell = cellIterator.next();
				String currentCellStr = currentCell.toString();
				switch (headerCell.toString()) {
				case "Award Organization Name":
					String deptName = currentCellStr.substring(currentCellStr.indexOf(" ") + 1, currentCellStr.length());
					if (!db.hasDepartment(deptName))
						db.addNewDepartment(deptName);
					currentDepartment = db.getDepartment(deptName);
					break;
				case "Award Principal Investigator":
					String personName = currentCellStr;
					if (!db.hasPerson(personName))
						db.addNewPerson(personName, currentDepartment);
					currentPerson = db.getPerson(personName);
					break;
				case "Current Budget":
					BigDecimal amount = new BigDecimal(currentCellStr);
					if (currentDepartment != null)
						currentDepartment.increaseAwardsTotal(amount);
					if (currentPerson != null)
						currentPerson.increaseAwardsTotal(amount);
				}
				headerCell = headerRow.getCell(++columnNum);
			}
			currentPerson = null;
			currentDepartment = null;
		}
		
	}
	
	public void write() {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		XSSFSheet departmentSheet = workbook.createSheet("Department Stats");
		XSSFSheet facultySheet = workbook.createSheet("Faculty Stats");
		
		XSSFCellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
		headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		headerStyle.setBorderBottom(BorderStyle.THIN);
		headerStyle.setBorderRight(BorderStyle.THIN);
		
		XSSFCellStyle dataStyle = workbook.createCellStyle();
		dataStyle.setBorderBottom(BorderStyle.THIN);
		dataStyle.setBorderRight(BorderStyle.THIN);
		
		Map<String, Object[]> departmentData = new TreeMap<String, Object[]>();
		Map<String, Object[]> facultyData = new TreeMap<String, Object[]>();
		departmentData.put("1", new Object[] {"DEPARTMENT", "PROPOSALS", "EXPENDITURES", "COMMITTED FUNDS", "STIPEND", "AWARDS"});
		facultyData.put("1", new Object[] {"PERSON", "DEPARTMENT", "PROPOSALS", "EXPENDITURES", "COMMITTED FUNDS", "STIPEND", "AWARDS"});
		
		Collection<Department> departmentValues = db.getDepartments();
		int lineCounter = 2;
		for (Department d : departmentValues) {
			departmentData.put(Integer.toString(lineCounter), new Object[] {d.getName(), d.getProposalsTotal(), d.getExpendituresTotal(), d.getCommittedFundsTotal(), d.getStipendTotal(), d.getAwardsTotal()});
			lineCounter++;
		}
		
		Set<String> keySet = departmentData.keySet();
		int rowNumber = 0;
		for (String key : keySet) {
			Row row = departmentSheet.createRow(rowNumber++);
			Object[] objArr = departmentData.get(key);
			int cellNumber = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellNumber++);
				
				if (rowNumber == 1) {
					cell.setCellStyle(headerStyle);
				} else {
					cell.setCellStyle(dataStyle);
				}
				
				if (obj instanceof String) {
					cell.setCellValue((String) obj);
				} else if (obj instanceof BigDecimal) {
					cell.setCellValue(currencyFormat((BigDecimal) obj));
				}
			}
		}
		
		
		
		Collection<Person> facultyValues = db.getFaculty();
		lineCounter = 2;
		for (Person p : facultyValues) {
			facultyData.put(Integer.toString(lineCounter), new Object[] {p.getName(), p.getDepartment(), p.getProposalsTotal(), p.getExpendituresTotal(), p.getCommittedFundsTotal(), p.getStipendTotal(), p.getAwardsTotal()});
			lineCounter++;
		}
		
		keySet = facultyData.keySet();
		rowNumber = 0;
		for (String key : keySet) {
			Row row = facultySheet.createRow(rowNumber++);
			Object[] objArr = facultyData.get(key);
			int cellNumber = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellNumber++);
				
				if (rowNumber == 1) {
					cell.setCellStyle(headerStyle);
				} else {
					cell.setCellStyle(dataStyle);
				}
				
				if (obj instanceof String) {
					cell.setCellValue((String) obj);
				} else if (obj instanceof Department) {
					cell.setCellValue(((Department) obj).getName());
				} else if (obj instanceof BigDecimal) {
					cell.setCellValue(currencyFormat((BigDecimal) obj));
				}
			}
		}
		
		
		// Resize the columns of the spreadsheets
		Row firstRow = departmentSheet.getRow(0);
		Iterator<Cell> firstRowIterator = firstRow.iterator();
		while (firstRowIterator.hasNext()) {
			Cell cell = firstRowIterator.next();
			departmentSheet.autoSizeColumn(cell.getColumnIndex());
		}
		firstRow = facultySheet.getRow(0);
		firstRowIterator = firstRow.iterator();
		while (firstRowIterator.hasNext()) {
			Cell cell = firstRowIterator.next();
			facultySheet.autoSizeColumn(cell.getColumnIndex());
		}
		
		// Finally, write the spreadsheet to output file
		FileOutputStream file = null;
		try {
			file = new FileOutputStream(outputFile);
			workbook.write(file);
			file.close();
			workbook.close();
		} catch (FileNotFoundException e) {
			System.out.println("FileNotFoundException: " + outputFile.toString());
			e.printStackTrace();
		} catch (IOException e) {
			System.out.println("Unable to write to " + outputFile.toString());
			e.printStackTrace();
		}
		
	}
	
	public static String currencyFormat(BigDecimal n) {
		return NumberFormat.getCurrencyInstance().format(n);
	}
	
}
