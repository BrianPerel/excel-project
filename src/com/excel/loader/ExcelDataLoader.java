package com.excel.loader;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.security.SecureRandom;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Properties;
import java.util.logging.Logger;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Main class - load configuration properties, select random names for the
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
	private static ArrayList<String> teamMembers = new ArrayList<>();
	private static final Logger logger_ = Logger.getLogger(ExcelDataLoader.class.getSimpleName());

	/**
	 * Adds data of names and dates to the array list to prepare to write to the
	 * excel file
	 * 
	 * @param random object
	 * @param dsuData data of names and dates
	 * @throws IOException
	 */
	private static void addData(ArrayList<MyData> dsuData) {
		SecureRandom random = new SecureRandom(LocalDateTime.now().toString().getBytes(StandardCharsets.US_ASCII));
		ArrayList<String> unusableNames = new ArrayList<>();

		// add empty row between headers and first data row
		dsuData.add(new MyData("", ""));

		for (int dayNumber = 0; dayNumber < daysInRotationSchedule; dayNumber++) {
			String name = teamMembers.get(random.nextInt(teamMembers.size())).trim();

			// while unusable names array list contains this name we have chosen, pick
			// another name. Array list is refreshed
			// for every week, if the day we just completed for was Friday
			while (unusableNames.contains(name) && !name.isEmpty() && teamMembers.size() > 5) {
				name = teamMembers.get(random.nextInt(teamMembers.size())).trim();
			}

			String date = getDsuDate(dayNumber);

			if (!isHoliday(dayNumber)) {
				// avoid adding Saturdays and Sundays to the schedule maker, by only adding the
				// record if the date doesn't
				// start with "Sat" or "Sun"
				if (!(date.startsWith("Sat") || date.startsWith("Sun"))) {
					dsuData.add(new MyData(name, date));
					unusableNames.add(name);
				}

				// add in an empty row to separate the weeks. If the date that was just added in
				// the loop started with "Fri"
				if (date.startsWith("Fri")) {
					dsuData.add(new MyData("", ""));
					unusableNames.clear();
				}
			} 
			else {
				dsuData.add(new MyData("Company Holiday", date));
			}
		}
	}

	/**
	 * Creates the excel cells with given data
	 * 
	 * @param dsuData the loaded records of team names and set of dates
	 * @param spreadsheet the excel spreadsheet
	 * @param workbook the excel workbook
	 */
	private static void createExcelSheet(ArrayList<MyData> dsuData, XSSFSheet spreadsheet, XSSFWorkbook workbook) {
		int rowid = 0;
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
	 * Writes the created/formatted data to the specified excel file
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
		}
		catch (IOException ex1) {
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
		return LocalDate.parse(startDate).plusDays(dayNumber).format(DateTimeFormatter.ofPattern("EEE, MM/dd/yy"));
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
		return Arrays.stream(companyHolidays).anyMatch(holiday -> holiday.trim().equals(LocalDate.parse(startDate).plusDays(dayNumber).toString()));
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

		companyHolidays = properties.getProperty("com.fy22.holidays", "").split(",");
		teamName = properties.getProperty("team.name", "");
		teamName = (!teamName.isBlank() ? teamName + " - " : teamName) + "DSU lead schedule";
		startDate = properties.getProperty("start.date", LocalDate.now().toString());
		fileSaveLocation = properties.getProperty("excel.file.save.location", "./Team DSU Schedule.xlsx").replace("/", "\\");
		isFileOpeningEnabled = Boolean.parseBoolean(properties.getProperty("excel.file.open.after.creation", "false"));
		String tMembers = properties.getProperty("team.members", "Person 1, Person 2, Person 3");

		if (tMembers != null) {
			teamMembers.addAll(Arrays.asList(tMembers.split(",")));
			Collections.shuffle(teamMembers);
		}

		if (!properties.getProperty("days.in.rotation", "5").isBlank()) {
			daysInRotationSchedule = Integer.parseInt(properties.getProperty("days.in.rotation", "5"));
		}

		return ensureNecessaryValues(tMembers);
	} 
	catch (IOException ex) {
		logger_.severe(ex.getMessage());
		return false;
	}
  }
  
	/**
	 * Check to ensure all parameters exist and contain values.
	 * 
	 * @param String tMembers is inspected for pertinent properties.
	 * @return boolean if all fields are populated.
	 */
	private static boolean ensureNecessaryValues(String tMembers) {
		return tMembers != null && !tMembers.isEmpty() && teamName != null && startDate != null && !startDate.isEmpty()
			&& teamMembers != null && fileSaveLocation != null && !fileSaveLocation.isEmpty()
			&& daysInRotationSchedule > 0;
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
				createExcelSheet(dsuData, workbook.createSheet(teamName), workbook);
				createExcelFile(workbook);
			} 
			else {
				logger_.severe("Error, Could not create excel file");
			}
		} 
		catch (IOException e) {
			logger_.severe("Unable to create workbook. XSSFWorkbook creation error. " + e.getMessage());
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
			row.createCell(x).setCellValue((x == 0) ? "Team Member" : "DSU Date");
			row.getCell(x).setCellStyle(getCellStyle(workbook));
		}
	}
}
