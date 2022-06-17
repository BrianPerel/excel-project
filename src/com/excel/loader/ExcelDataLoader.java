package com.excel.loader;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.Properties;
import java.util.Stack;

import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Main class - loads configuration properties, selects random names for the
 * dates, creates an excel workbook and spreadsheet, formats the data, and adds
 * it to the excel file.
 * 
 * @author Brian Perel
 *
 */
public class ExcelDataLoader {
	
	private static String teamName;
	private static String startDate;
	private static String fileSaveLocation;
	private static String[] companyHolidays;
	private static int rotationDaysSchedule;
	private static boolean isFileOpeningEnabled;
	private static final String DATE_FORMAT = "MM/dd/yyyy";
	private static Stack<String> teamMembersStack = new Stack<>();
	private static ArrayList<String> teamMembers = new ArrayList<>();
	private static final Logger logger_ = Logger.getLogger(ExcelDataLoader.class);

	/**
	 * Adds data of names and dates to the array list to prepare to write to the
	 * excel file
	 * 
	 * @param random object
	 * @param dsuData data list of names and dates
	 * @throws IOException
	 */
	private static void addData(ArrayList<MyData> dsuData) {
		teamMembersStack = getStack(teamMembers); // create stack before looping
		
		dsuData.add(new MyData()); // add an empty row to separate headers row and 1st data row

		for (int dayNumber = 0; dayNumber < rotationDaysSchedule; dayNumber++) {
			String date = getDsuDate(dayNumber);
			
			if(isHoliday(dayNumber)) {
				dsuData.add(new MyData("Company Holiday", date));
			}
			else {
				// avoid adding Saturdays and Sundays to the schedule maker
				if (StringUtils.startsWithIgnoreCase(date, "Sat") || StringUtils.startsWithIgnoreCase(date, "Sun")) {
					continue;
				}
	
				dsuData.add(new MyData(getNextName(teamMembersStack).trim(), date));
				
				if (StringUtils.startsWithIgnoreCase(date, "Fri")) {
					dsuData.add(new MyData());
				} 
			}
		}
	}

	/**
	 * Gets next name for DSU rotation schedule while enforcing that every person goes exactly once before they can go a second time
	 * in the full rotation schedule. Pop and use the next name from the stack. 
	 * If stack is empty create a new stack from the full team member list
	 * 
	 * Using a stack to ensure that every person in the shuffled list goes at least once in the full rotation, avoid duplicates,
	 * and get a name while removing it during the pop operation
	 * 
	 * @param argTeamMembersStack shuffled stack of random names
	 * @return next name from the stack of names
	 */
	private static String getNextName(Stack<String> argTeamMembersStack) {
		// using recursion to populate the stack if needed and then immediately calling it again to perform the requested action
		if(argTeamMembersStack.isEmpty()) {
			argTeamMembersStack.addAll(getStack(teamMembers));
			return getNextName(argTeamMembersStack);
		}
		
		return argTeamMembersStack.pop();
	}

	/**
	 * Creates and returns a shuffled stack of names 
	 * 
	 * @param argTeamMembers the names loaded from the configuration file
	 * @return shuffled stack of names
	 */
	private static Stack<String> getStack(ArrayList<String> argTeamMembers) {
		Stack<String> stack = new Stack<>();
		
		// create roster of names
		List<String> members = new ArrayList<>(argTeamMembers);
		Collections.shuffle(members); // add the random aspect by shuffling the names loaded from the configuration file
		
		stack.addAll(members);
		return stack;
	}

	/**
	 * Creates the excel cells with given data
	 * 
	 * @param dsuData the loaded records of team names and set of dates
	 * @param spreadsheet the excel spreadsheet
	 * @param workbook the excel workbook
	 */
	private static void createExcelSheet(ArrayList<MyData> dsuData, XSSFWorkbook workbook) {
		int rowid = 0;
		XSSFSheet spreadsheet = workbook.createSheet(teamName + "DSU lead schedule");
		XSSFRow row = spreadsheet.createRow(rowid++);

		setHeaders(workbook, row);

		// writing the data into the sheets
		for (MyData dataRecord : dsuData) {
			row = spreadsheet.createRow(rowid++);
			row.createCell(0).setCellValue(getCellValue(dataRecord, 0));
			row.createCell(1).setCellValue(getCellValue(dataRecord, 1));
		}
	}

	/**
	 * Writes the formatted data to the specified excel file
	 * 
	 * @param workbook the excel workbook
	 */
	private static void createExcelFile(XSSFWorkbook workbook) {
		try (FileOutputStream out = new FileOutputStream(new File(fileSaveLocation))) {
			workbook.write(out);
			logger_.info("Process completed. Excel file created");

			if (isFileOpeningEnabled) {
				logger_.info("File opening feature is enabled. Opening " + fileSaveLocation + " file...");
				Desktop.getDesktop().open(new File(fileSaveLocation));
			}
		} catch (IOException ex1) {
			logger_.error(ex1.getMessage());
		}
	}

	/**
	 * Creates custom excel styling options
	 * 
	 * @param workbook the excel workbook
	 * @return the custom style
	 */
	private static XSSFCellStyle getCellStyle(XSSFWorkbook workbook) {
		XSSFCellStyle style = workbook.createCellStyle(); // create style
		XSSFFont font = workbook.createFont(); // create font
		font.setBold(true); // set to bold font
		font.setFontHeight(12); // set font height
		style.setFont(font); // set font
		return style;
	}

	/**
	 * Determines and gets data to be placed in the appropriate column
	 * 
	 * @param dataRecord current record being retrieved from the data set
	 * @param cellid id of the cell currently being targeted
	 * @return data going into the excel sheet
	 */
	private static String getCellValue(MyData dataRecord, int cellid) {
		return (cellid == 0) ? dataRecord.getName() : dataRecord.getDate();
	}

	/**
	 * Gets the start date set in configuration file, updates it by the day that it
	 * is in rotation, formats it as wanted
	 * 
	 * @param dayNumber the day number in the rotation schedule
	 * @return formatted day for current record to be added to array list
	 */
	private static String getDsuDate(int dayNumber) {
		return LocalDate.parse(startDate, DateTimeFormatter.ofPattern(DATE_FORMAT)).plusDays(dayNumber)
				.format(DateTimeFormatter.ofPattern("EEE, " + DATE_FORMAT));
	}

	/**
	 * Checks to see if current date is a company holiday date
	 * 
	 * @param dsuData data set that we're adding data to
	 * @param dayNumber the day number in the rotation schedule
	 * @param date current date to analyze
	 * @return if it's a holiday
	 */
	private static boolean isHoliday(int dayNumber) {
		return Arrays.stream(companyHolidays)
				.anyMatch(holiday -> holiday.trim()
				.equals(LocalDate.parse(startDate, DateTimeFormatter.ofPattern(DATE_FORMAT)).plusDays(dayNumber)
				.format(DateTimeFormatter.ofPattern(DATE_FORMAT))));
	}

	/**
	 * Loads configurable properties that will be included in the excel sheet
	 * 
	 * @return boolean value determining if the properties were able to be loaded
	 */
	private static boolean loadConfigurations() {
		Properties properties = new Properties();

		try (InputStream input = new FileInputStream("excel-sheet.properties")) {
			properties.load(input);

			companyHolidays = properties.getProperty("com.us.fy22.holidays", "").split(",");
			teamName = properties.getProperty("team.name", "Team Orion") + " ";
			startDate = properties.getProperty("start.date", LocalDate.now().format(DateTimeFormatter.ofPattern(DATE_FORMAT)));
			fileSaveLocation = properties.getProperty("excel.file.save.location", "Team DSU Schedule.xlsx").replace("/", "\\");
			
			if(!StringUtils.endsWithIgnoreCase(fileSaveLocation, ".xlsx")) {
				fileSaveLocation = fileSaveLocation.replace(fileSaveLocation.substring(fileSaveLocation.indexOf('.')), ".xlsx");
			}
			
			isFileOpeningEnabled = Boolean.parseBoolean(properties.getProperty("excel.file.opening.enabled", "false"));
			rotationDaysSchedule = StringUtils.isNumeric(properties.getProperty("rotation.days", "5")) ?
					Integer.parseInt(properties.getProperty("rotation.days", "5")) : 5;

			return ensureNecessaryValues(properties);
			
		} catch (IOException ex) {
			logger_.error(ex.getMessage());
			return false;
		}
	}

	/**
	 * Validates that all the loaded configuration file values contain something
	 * 
	 * @param tMembers team member's names
	 * @return pass or fail boolean value
	 */
	private static boolean ensureNecessaryValues(Properties properties) {
		if(startDate.isEmpty()) {
			startDate = LocalDate.now().format(DateTimeFormatter.ofPattern(DATE_FORMAT));
		}
		
		String tMembers = properties.getProperty("team.members", "");
		
		if (tMembers != null) {
			teamMembers.addAll(Arrays.asList(tMembers.split(",")));
		}
		
		for(int x = 0; x < teamMembers.size(); x++) {
			// Capitalize both first and last names of every record in names list
			teamMembers.set(x, teamMembers.get(x).trim());
			teamMembers.set(x, (StringUtils.capitalize(teamMembers.get(x).substring(0, teamMembers.get(x).indexOf(' ') + 1))
					+ StringUtils.capitalize(teamMembers.get(x).substring(teamMembers.get(x).indexOf(' ') + 1))));
		}
		
		if (rotationDaysSchedule > 1000) {
			logger_.warn("Max days allowed in rotation schedule exceeded. Setting number to 100 days for rotation");
			rotationDaysSchedule = 100;
		}
		
		return companyHolidays != null && tMembers != null && teamName != null && startDate != null && teamMembersStack != null
				&& fileSaveLocation != null;
	}

	/**
	 * Main method for executing the application
	 *
	 * @param args from command line
	 */
	public static void main(String... args) {
		System.setProperty("logs.location", "ExcelDataLoader.log");
		PropertyConfigurator.configure(ExcelDataLoader.class.getResourceAsStream("log4j.properties"));
		logger_.info("Application Started...");
		
		try (XSSFWorkbook workbook = new XSSFWorkbook()) {
			if (loadConfigurations()) {
				Files.deleteIfExists(new File(fileSaveLocation).toPath());
				ArrayList<MyData> dsuData = new ArrayList<>();
				addData(dsuData);
				createExcelSheet(dsuData, workbook);
				createExcelFile(workbook);
			} else {
				logger_.error("Error, Could not create excel file");
			}
		} catch (IOException e) {
			logger_.error("Unable to create workbook. XSSFWorkbook creation error. " + e.getMessage());
		}
	}

	/**
	 * Sets the excel table's header to the specified font style
	 * 
	 * @param workbook the excel workbook
	 * @param row the excel sheet row
	 */
	private static void setHeaders(XSSFWorkbook workbook, XSSFRow row) {
		// format header rows to be bold and slightly larger size
		for (int x = 0; x < 2; x++) {
			row.createCell(x).setCellValue(x == 0 ? "Team Member" : "DSU Date");
			row.getCell(x).setCellStyle(getCellStyle(workbook));
		}
	}
}
