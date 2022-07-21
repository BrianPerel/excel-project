package com.excel.loader;

import static org.apache.poi.ss.usermodel.IndexedColors.GREY_25_PERCENT;
import static org.apache.poi.ss.usermodel.IndexedColors.LIGHT_YELLOW;
import static org.apache.poi.ss.usermodel.IndexedColors.SKY_BLUE;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
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
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Main class - loads properties from a configuration file, selects random names for the
 * dates, creates an excel workbook and spreadsheet, formats the data, and adds
 * it to the excel file.
 * 
 * @author Brian Perel
 * @created April 22, 2022
 */
public class DsuLeadGenerator {
	
	private static final String DATE_FORMAT = "MM/dd/yyyy";
	private static final Logger logger_ = Logger.getLogger(DsuLeadGenerator.class);
	
	private String teamName;
	private String startDate;
	private String fileSaveLocation;
	private String[] companyHolidays;
	private int rotationDaysSchedule;
	private boolean isFileOpeningEnabled;
	private Stack<String> teamMembersStack = new Stack<>();
	private List<String> teamMembers = new ArrayList<>(10);

	/**
	 * Adds data of names and dates to the array list to prepare to write to the excel file
	 * 
	 * @param random object
	 * @param dsuData data list of names and dates
	 * @throws IOException
	 */
	private void addData(List<MyData> dsuData) {
		List<String> unusableNames = new ArrayList<>(5);
		teamMembersStack = getStack(); // create stack before looping
		dsuData.add(new MyData()); // add an empty row to separate headers row and 1st data row

		for (int dayNumber = 0; dayNumber < rotationDaysSchedule; dayNumber++) {
			String date = getDsuDate(dayNumber);
			
			if(isHoliday(dayNumber)) {
				dsuData.add(new MyData("Company Holiday", date));
			}
			else {
				if (StringUtils.startsWithIgnoreCase(date, "Sat") || StringUtils.startsWithIgnoreCase(date, "Sun")) {
					continue; // avoid adding Saturdays and Sundays to the schedule maker
				}
				
				String chosenName = avoidDuplicateNames(unusableNames); 
				
				// avoids an infinite loop by avoiding duplicate name comparison, if number of team members entered is 5 or less. 
				// Since we have 5 days in a week and having 5 or less people wouldn't allow us to avoid duplicates
				if(teamMembers.size() > 5) {
					unusableNames.add(chosenName);
				}
	
				dsuData.add(new MyData(chosenName, date));
				
				if (StringUtils.startsWithIgnoreCase(date, "Fri")) {
					dsuData.add(new MyData());
					unusableNames.clear();
				} 
			}
		}
	}

	/**
	 * Prevents duplicate names from being used in any 1 week
	 * 
	 * @param unusableNames names that have already been used this week
	 * @return a uniquely chosen name 
	 */
	private String avoidDuplicateNames(List<String> unusableNames) {
		String name = getNextName(teamMembersStack);
		
		while(unusableNames.contains(name)) {
			name = getNextName(teamMembersStack);
		}
		
		return name.trim();
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
	private String getNextName(Stack<String> argTeamMembersStack) {
		// populate the stack if it's empty
		if(argTeamMembersStack.isEmpty()) {
			argTeamMembersStack.addAll(getStack());
		}
		
		return argTeamMembersStack.pop();
	}

	/**
	 * Creates and returns a shuffled stack of names 
	 * 
	 * @param argTeamMembers the names loaded from the configuration file
	 * @return shuffled stack of names
	 */
	private Stack<String> getStack() {
		Stack<String> stack = new Stack<>();
		
		Collections.shuffle(teamMembers); // add the random aspect by shuffling the names loaded from the configuration file
		
		stack.addAll(teamMembers);
		return stack;
	}

	/**
	 * Creates the excel cells with given data
	 * 
	 * @param dsuData the loaded records of team names and set of dates
	 * @param spreadsheet the excel spreadsheet
	 * @param workbook the excel workbook
	 */
	private void createExcelSheet(List<MyData> dsuData, XSSFWorkbook workbook) {
		int rowid = 0;
		XSSFSheet spreadsheet = workbook.createSheet(teamName.concat("DSU lead schedule"));
		XSSFRow row = spreadsheet.createRow(rowid++);

		setHeaders(workbook, row);

		// writing the data into the excel sheet
		for (MyData dataRecord : dsuData) {
			row = spreadsheet.createRow(rowid++);
			row.createCell(0).setCellValue(getCellValue(dataRecord, 0));
			row.createCell(1).setCellValue(getCellValue(dataRecord, 1));
			
			// if row is not apply gray color background
			if(!dataRecord.getName().isEmpty()) {
				row.getCell(0).setCellStyle(getCellStyle(workbook, GREY_25_PERCENT, false));
				row.getCell(1).setCellStyle(getCellStyle(workbook, GREY_25_PERCENT, false));
			}
			
			// if company holiday apply yellow color background to cell
			if(row.getCell(0).getStringCellValue().equalsIgnoreCase("Company Holiday")) {
				row.getCell(0).setCellStyle(getCellStyle(workbook, LIGHT_YELLOW, true));
				row.getCell(1).setCellStyle(getCellStyle(workbook, LIGHT_YELLOW, true));
			}
		}
		
		// permanently auto expands selected cell 
		spreadsheet.autoSizeColumn(0);
		spreadsheet.autoSizeColumn(1);
	}

	/**
	 * Writes the formatted data to the specified excel file
	 * 
	 * @param workbook the excel workbook
	 */
	private void createExcelFile(XSSFWorkbook workbook) {
		try (FileOutputStream out = new FileOutputStream(new File(fileSaveLocation))) {
			logger_.info("Process completed. Excel file created");
			workbook.write(out);

			if (isFileOpeningEnabled) {
				logger_.info("File opening feature is enabled. Opening \'" + fileSaveLocation + "\' file...");
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
	private XSSFCellStyle getCellStyle(XSSFWorkbook workbook, IndexedColors color, boolean isCustomFontNeeded) {
		XSSFCellStyle style = workbook.createCellStyle(); // create style
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND); // set the fill pattern to be solid whole
		style.setFillForegroundColor(color.index); // set the fill color style
		style.setBorderBottom(BorderStyle.THIN); // create border styles
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		
		// custom font is only needed for headers and holiday date cells
		if(isCustomFontNeeded) {
			XSSFFont font = workbook.createFont(); // create font
			font.setBold(true); // set to bold font
			font.setFontHeight(12); // set font height
			style.setFont(font); // set font
		}
		
		return style;
	}

	/**
	 * Determines and gets data to be placed in the appropriate column
	 * 
	 * @param dataRecord current record being retrieved from the data set
	 * @param cellid id of the cell currently being targeted
	 * @return data going into the excel sheet
	 */
	private String getCellValue(MyData dataRecord, int cellid) {
		return (cellid == 0) ? dataRecord.getName() : dataRecord.getDate();
	}

	/**
	 * Gets the start date set in configuration file, updates it by the day that it
	 * is in rotation, formats it as wanted
	 * 
	 * @param dayNumber the day number in the rotation schedule
	 * @return formatted day for current record to be added to array list
	 */
	private String getDsuDate(int dayNumber) {
		return LocalDate.parse(startDate, DateTimeFormatter.ofPattern(DATE_FORMAT)).plusDays(dayNumber)
				.format(DateTimeFormatter.ofPattern("EEEE, " + DATE_FORMAT));
	}

	/**
	 * Checks to see if current date is a company holiday date
	 * 
	 * @param dsuData data set that we're adding data to
	 * @param dayNumber the day number in the rotation schedule
	 * @param date current date to analyze
	 * @return if it's a holiday
	 */
	private boolean isHoliday(int dayNumber) {
		return Arrays.stream(companyHolidays)
				.anyMatch(holiday -> holiday.trim()
				.equals(getDsuDate(dayNumber).substring(getDsuDate(dayNumber).indexOf(' ') + 1)));
	}

	/**
	 * Loads configurable properties that will be included in the excel sheet
	 * 
	 * @return boolean value determining if the properties were able to be loaded
	 */
	private boolean loadConfigurations(String arg) {
		Properties table = new Properties();

		try (InputStream input = new FileInputStream(arg)) {
			table.load(input);

			companyHolidays = table.getProperty("com.us.fy22.holidays", "").split(",");
			teamName = table.getProperty("team.name", "Team Orion") + " ";
			startDate = table.getProperty("start.date", LocalDate.now().format(DateTimeFormatter.ofPattern(DATE_FORMAT)));
			fileSaveLocation = table.getProperty("excel.file.save.location", "Team DSU Schedule.xlsx").replace("/", "\\");
			
			if(!StringUtils.endsWithIgnoreCase(fileSaveLocation, ".xlsx")) {
				logger_.warn("Provided excel file name is of incorrect file extension. Setting to .xlsx for excel");
				fileSaveLocation = fileSaveLocation.replace(fileSaveLocation.substring(fileSaveLocation.indexOf('.')), ".xlsx");
			}
			
			isFileOpeningEnabled = Boolean.parseBoolean(table.getProperty("excel.file.opening.enabled", "false"));
			rotationDaysSchedule = (StringUtils.isNumeric(table.getProperty("rotation.days", "5"))) ?
					Integer.parseInt(table.getProperty("rotation.days", "5")) : 5;

			return ensureNecessaryValues(table);
			
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
	private boolean ensureNecessaryValues(Properties properties) {
		if(startDate.length() != 10 || startDate.matches("[a-zA-Z]+\\.?")) {
			String currentDate = LocalDate.now().format(DateTimeFormatter.ofPattern(DATE_FORMAT));
			logger_.warn("Provided start date is empty or of incorrect format. Setting start date to current date: \'" + currentDate + "\'");
			startDate = currentDate;
		}
		
		String tMembers = properties.getProperty("team.members", "Person 1, Person 2, Person 3, Person 4, Person 5");

		if (tMembers != null) {
			teamMembers.addAll(Arrays.asList(tMembers.split(","))); // create roster of names
			
			if(teamMembers.size() == 1 && teamMembers.get(0).equals("")) {
				teamMembers.remove(0);
				teamMembers.addAll(Arrays.asList(("Person 1, Person 2, Person 3, Person 4, Person 5").split(",")));
				logger_.warn("Team members property list was empty. Adding elements to list");
			}
			else if(teamMembers.size() < 5) {
				int membersNeeded = 5 - teamMembers.size();
				
				for(int x = 0; x < membersNeeded; x++) {
					int y = x;
					teamMembers.add("Person " + (++y));
				}
				
				logger_.warn("team members property list was of a lesser size than required, additional element names have been added");
			}
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
		PropertyConfigurator.configure("log4j.properties");
		logger_.info("Application Started...");
		
		DsuLeadGenerator run = new DsuLeadGenerator();
		run.execute(args.length == 0 ? "excel-sheet.properties" : args[0]);
	}

	private void execute(String arg) {
		try (XSSFWorkbook workbook = new XSSFWorkbook()) {
			boolean isLoadSuccessful = loadConfigurations(arg);
			
			if (isLoadSuccessful) {
				List<MyData> dsuData = new ArrayList<>();
				addData(dsuData);
				createExcelSheet(dsuData, workbook);
				createExcelFile(workbook);
			} else {
				logger_.error("Error, Could not load excel file property settings. Please check your \'excel-sheet.properties\' configurations");
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
	private void setHeaders(XSSFWorkbook workbook, XSSFRow row) {
		// format header rows to be bold and slightly larger size
		for (int x = 0; x < 2; x++) {
			row.createCell(x).setCellValue(x == 0 ? "Team Member" : "DSU Date");
			row.getCell(x).setCellStyle(getCellStyle(workbook, SKY_BLUE, true));
		}
	}
}
