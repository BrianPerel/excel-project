package com.excel.loader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.security.SecureRandom;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.Properties;
import java.util.logging.Logger;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataLoader {

  private static String teamName;
  private static String startDate;
  private static String fileSaveLocation;
  private static int daysInRotationSchedule;
  private static String[] teamMembers = new String[9];
  private static final Logger logger_ = Logger.getLogger(ExcelDataLoader.class.getSimpleName());

  private static void addData(SecureRandom random, ArrayList<MyData> dsuData) {
    // TODO: prevent duplicate names from being added in the same week
    for (int dayNumber = 0; dayNumber < daysInRotationSchedule; dayNumber++ ) {
      String name = teamMembers[random.nextInt(teamMembers.length)];
      String date = getDsuDate(dayNumber);

      // avoid adding Saturdays and Sundays to the schedule maker, by only adding the record if the date doesn't
      // start with "Sat" or "Sun"
      if (!date.startsWith("Sat") && !date.startsWith("Sun")) {
        dsuData.add(new MyData(name, date));
      }

      // add in an empty row to separate the weeks. If the date that was just added in the loop started with "Fri"
      if (date.startsWith("Fri")) {
        dsuData.add(new MyData("", ""));
      }
    }
  }

  /**
   * Creates the excel cells with given data
   * @param dsuData the loaded records of team names and set of dates
   * @param spreadsheet the excel spreadsheet
   * @param workbook the excel workbook
   */
  private static void createExcelSheet(ArrayList<MyData> dsuData, XSSFSheet spreadsheet, XSSFWorkbook workbook) {
    int rowid = 0;
    XSSFRow row = spreadsheet.createRow(rowid++ );

    setHeaders(workbook, row);

    // writing the data into the sheets...
    for (MyData record : dsuData) {
      row = spreadsheet.createRow(rowid++ );
      row.createCell(0).setCellValue(getCellValue(record, 0));
      row.createCell(1).setCellValue(getCellValue(record, 1));
    }

    createFile(workbook);
  }

  /**
   * Writes the created/formatted data to the specified excel file
   * @param workbook the excel workbook
   */
  private static void createFile(XSSFWorkbook workbook) {
    try (FileOutputStream out = new FileOutputStream(new File(fileSaveLocation))) {
      workbook.write(out);
      logger_.info("Process complete - excel file created");
    }
    catch (IOException ex1) {
      logger_.severe(ex1.getMessage());
    }
  }

  /**
   * Creates custom excel styling options
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
   * @param record current record being retrieved from the data set
   * @param cellid id of the cell currently being targeted
   * @return data going into the excel sheet
   */
  private static String getCellValue(MyData record, int cellid) {
    return (cellid == 0) ? record.getName() : record.getDate();
  }

  @SuppressWarnings("deprecation")
  private static String getDsuDate(int dayNumber) {
    return new SimpleDateFormat("EEE, MM/dd/yy")
        .format(new Date(startDate).getTime() + (1000 * 60 * 60 * dayNumber * 24));
  }

  /**
   * Loads configurable properties that will be included in the excel sheet
   * @return boolean value determining if the properties were able to be loaded or not
   */
  private static boolean loadConfigurations() {
    Properties properties = new Properties();

    try (InputStream input = new FileInputStream("excel_sheet.properties")) {
      properties.load(input);

      teamName = properties.getProperty("team.name");
      startDate = properties.getProperty("start.date");
      teamMembers = properties.getProperty("team.members").split(",");
      fileSaveLocation = properties.getProperty("excel.file.save.location");
      daysInRotationSchedule = Integer.parseInt(properties.getProperty("days.in.rotation"));
      return true;
    }
    catch (IOException ex) {
      logger_.severe(ex.getMessage());
      return false;
    }
  }

  public static void main(String[] args) {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      SecureRandom random = new SecureRandom(LocalDateTime.now().toString().getBytes(StandardCharsets.US_ASCII));
      
      ArrayList<MyData> dsuData = new ArrayList<>();

      if (loadConfigurations()) {
        // removes extra space after obtaining team member names
        for (int index = 0; index < teamMembers.length; index++ ) {
          teamMembers[index] = teamMembers[index].trim();
        }

        addData(random, dsuData);

        createExcelSheet(dsuData, workbook.createSheet(teamName + " - DSU lead schedule"), workbook);
      }
    }
    catch (IOException e) {
      logger_.severe("Unable to create workbook. XSSFWorkbook creation error. " + e.getMessage());
    }
  }

  /**
   * Sets the excel table's header to the specified font style
   * @param workbook the excel workbook
   * @param row the excel sheet row
   */
  private static void setHeaders(XSSFWorkbook workbook, XSSFRow row) {
    XSSFCellStyle style = getCellStyle(workbook);

    // format header rows to be bold and sligtly higher size
    row.createCell(0).setCellValue("Team Member");
    row.getCell(0).setCellStyle(style);
    row.createCell(1).setCellValue("DSU Date");
    row.getCell(1).setCellStyle(style);
  }
}
