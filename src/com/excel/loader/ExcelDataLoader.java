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
import java.util.Collection;
import java.util.Collections;
import java.util.List;
import java.util.Properties;
import java.util.Stack;
import java.util.logging.Logger;

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
	private static int daysInRotationSchedule;
	private static boolean isFileOpeningEnabled;
	private static final String DATE_FORMAT = "MM/dd/yyyy";
	private static Stack<String> teamMembersStack = new Stack<>();
	private static ArrayList<String> teamMembers = new ArrayList<>();
	private static final Logger logger_ = Logger.getLogger(ExcelDataLoader.class.getSimpleName());

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
		// add an empty row to separate headers row and 1st data row
		dsuData.add(new MyData());

		for (int dayNumber = 0; dayNumber < daysInRotationSchedule; dayNumber++) {
			String date = getDsuDate(dayNumber);

			if(isHoliday(dayNumber)) {
				dsuData.add(new MyData("Company Holiday", date));
			}
			else {
				// avoid adding Saturdays and Sundays to the schedule maker
				if (date.startsWith("Sat") || date.startsWith("Sun")) {
					continue;
				}
	
				dsuData.add(new MyData(getNextName(teamMembersStack).trim(), date));
				
				if (date.startsWith("Fri")) {
					dsuData.add(new MyData());
				} 
			}
		}
	}

	/**
	 * Gets next name for DSU rotation schedule while enforcing that every person goes exactly once before they can go a second time
	 * in the full rotation schedule. Pop and use the next name from the stack. 
	 * If stack is empty create a new stack from the full team member list. 
	 * 
	 * Using a stack to ensure that every person in the shuffled list goes at least once in the full rotation, avoid duplicates,
	 * and get a name while removing it during the pop operation
	 * 
	 * @param argTeamMembersStack shuffled stack of random names
	 * @return next name from the stack of names
	 */
	private static String getNextName(Stack<String> argTeamMembersStack) {
		// using recursion to populate stack and then calling it again to perform the requested action
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
	private static Stack<String> getStack(Collection<String> argTeamMembers) {
		Stack<String> stack = new Stack<>();
		
		// create roster of names
		List<String> members = new ArrayList<>(argTeamMembers);
		Collections.shuffle(members); // add the random aspect by shuffling the names loaded from the config
		
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
			logger_.info("Process complete - excel file created");

			if (isFileOpeningEnabled) {
				Desktop.getDesktop().open(new File(fileSaveLocation));
			}
		} catch (IOException ex1) {
			logger_.severe(ex1.getMessage());
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
	 * @param dsuData data list that we're adding data to
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
			isFileOpeningEnabled = Boolean.parseBoolean(properties.getProperty("excel.file.opening.enabled", "false"));
			daysInRotationSchedule = (properties.getProperty("rotation.days", "5").length() != 0) ?
					Integer.parseInt(properties.getProperty("rotation.days", "5")) : 5;
			
			String tMembers = properties.getProperty("team.members", "Person 1, Person 2");
			return ensureNecessaryValues(tMembers);
			
		} catch (IOException ex) {
			logger_.severe(ex.getMessage());
			return false;
		}
	}

	/**
	 * Validates that all the loaded config file values contain something
	 * 
	 * @param tMembers team member names
	 * @return pass or fail boolean value
	 */
	private static boolean ensureNecessaryValues(String tMembers) {
		if(startDate.isBlank())	{
			startDate = LocalDate.now().format(DateTimeFormatter.ofPattern(DATE_FORMAT));
		}
		
		if (tMembers != null) {
			teamMembers.addAll(Arrays.asList(tMembers.split(",")));
		}
		
		if (daysInRotationSchedule > 1000) {
			daysInRotationSchedule = 100;
		}
		
		return tMembers != null && teamName != null && startDate != null && teamMembersStack != null
				&& fileSaveLocation != null;
	}

	/**
	 * Main method for executing the application
	 * 
	 * @param args from command line
	 */
	public static void main(String... args) {
		try (XSSFWorkbook workbook = new XSSFWorkbook()) {
			if (loadConfigurations()) {
				Files.deleteIfExists(new File(fileSaveLocation).toPath());
				ArrayList<MyData> dsuData = new ArrayList<>();
				addData(dsuData);
				createExcelSheet(dsuData, workbook);
				createExcelFile(workbook);
			} else {
				logger_.severe("Error, Could not create excel file");
			}
		} catch (IOException e) {
			logger_.severe("Unable to create workbook. XSSFWorkbook creation error. " + e.getMessage());
		}
	}

	/**
	 * Sets the excel table's headers to the specified font style
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
