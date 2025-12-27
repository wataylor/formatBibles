package asst.hssf;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;

/**
 * Class to hold information about the work book being read.  Workbook
 * and Sheet are interfaces which can hold implementations for both.
 * The plan is to be able to read either type of spread sheet and
 * process them in the same way using the same libraries.  The goal is
 * to be able to handle both xls and xlsx documents.
 * @author Material Gain
 * @2015 01
 */

public class WorkbookManager {
  /** Column map in the current sheet*/
  public Map<String, Integer> columnMap;
  /** The current row being processed*/
  public Row row;
  /** The current work sheet in the work book. */
  public Sheet sheet;
  /** Name of the current sheet */
  public String sheetName;
  /** Error accumulator for all work sheet operations. */
  public StringBuilder sb = new StringBuilder();
  /** This stores a workbook that came either from .xls or .xlsx, but it
   * remembers which it was and writes itself accordingly.*/
  public Workbook wb;
  public CreationHelper createHelper;
  public CellStyle cellDateStyle;
  public CellStyle cellDateTimeStyle;
  public CellStyle cellCurrencyFormat;
  public CellStyle headerTextStyle;
  public CellStyle integerSty;
  public CellStyle float2Sty;
  /** Store the name of the file which is associated with the work book*/
  public String fileName;
  /** True means that the work book has to be written to a new file*/
  public boolean dirty = false;
  /** True means to update the database if all rows are good */
  public boolean doIt;
  /** True means to generate extra sysout. */
  public boolean verbose;
  /** True means that the work sheet is updating existing rows, false means
   * that the work sheet is creating new rows from spread sheet rows*/
  public boolean updating;
  /** Database connection manager.  It is up to the caller to close the
   * database.  This field requires many database resources.  */
  //public ConnData dbd;

  /** The 0-based index of the current row being processed*/
  public int currentRowNumber;
  protected int firstRowNumber;
  protected int lastRowNumber;
  /** Number of actual rows in the spread sheet*/
  protected int physicalRows;
  protected int sheetIndex;

  /**
   * Default constructor
   */
  public WorkbookManager() { /* */ }

  /**
   * Constructor which grabs the first sheet in a work book
   * @param wb the workbook of either kind
   */
  public WorkbookManager(Workbook wb) {
    this.wb = wb;
    createHelper = wb.getCreationHelper();
    sheet = wb.getSheetAt(0);
    sheetCharacteristics();
    makeFormatters();
  }

  /**
   * Constructor which selects or creates a named sheet if asked to do
   * so
   * @param wb the workbook of either kind
   * @param sheetName the name of the desired sheet
   * @param create tells whether to create the sheet if it does not exist
   */
  public WorkbookManager(Workbook wb, String sheetName, boolean create) {
    this.wb = wb;
    createHelper = wb.getCreationHelper();
    if (( (sheet = wb.getSheet(sheetName)) == null) && create) {
      sheet = wb.createSheet(sheetName);
    }
    sheetCharacteristics();
    makeFormatters();
  }

  /**
   * Constructor which selects or creates a numbered sheet and creates
   * one new sheet if asked to do so
   * @param wb the workbook of either kind
   * @param sheetNum the 0-based number of the desired sheet
   * @param create tells whether to create the sheet if it does not exist
   */
  public WorkbookManager(Workbook wb, int sheetNum, boolean create) {
    this.wb = wb;
    createHelper = wb.getCreationHelper();
    if (( (sheet = wb.getSheetAt(sheetNum)) == null) && create) {
      sheet = wb.createSheet();
    }
    sheetCharacteristics();
    makeFormatters();
  }

  /** Create a number of useful cell formatters */
  public void makeFormatters() {
    DataFormat format = wb.createDataFormat();
    createHelper = wb.getCreationHelper();
    cellDateStyle = wb.createCellStyle();
    cellDateStyle.setDataFormat(
	createHelper.createDataFormat().getFormat("m/d/yy"));
    cellDateTimeStyle = wb.createCellStyle();
    cellDateTimeStyle.setDataFormat(
	createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
    cellCurrencyFormat = wb.createCellStyle();
    cellCurrencyFormat.setDataFormat((short)7);
    headerTextStyle = wb.createCellStyle();
    XSSFFont txtFont = (XSSFFont) wb.createFont();
    txtFont.setFontHeightInPoints((short)14);
    txtFont.setBold(true);
    headerTextStyle.setFont(txtFont);
    integerSty = wb.createCellStyle();
    integerSty.setDataFormat(format.getFormat("#"));
    float2Sty = wb.createCellStyle();
    float2Sty.setDataFormat(format.getFormat("0.0"));
  }

  /**
   * Selects a specified sheet as the current sheet.  Leaves sheet null
   * if there is no sheet by that name.
   * @param desiredSheetName name of the desired sheet.
   * @return The desired sheet or null.
   */
  public Sheet pickSheet(String desiredSheetName) {
    if ( (sheet = wb.getSheet(desiredSheetName)) == null) { return null; }
    sheetCharacteristics();
    return sheet;
  }

  /**
   * Compute various interesting characteristics of the specified
   * sheet within the work book
   */
  public void sheetCharacteristics() {
    if (sheet == null) { return; }
    sheetName = sheet.getSheetName();
    lastRowNumber = sheet.getLastRowNum();
    firstRowNumber = sheet.getFirstRowNum();
    physicalRows = sheet.getPhysicalNumberOfRows();
    if (physicalRows != 0) {
      row = sheet.getRow(firstRowNumber);
      currentRowNumber = firstRowNumber;
    } else {
      currentRowNumber = -1;
      row = null;
    }
    // sheet.createFreezePane( 0, 0 );  // gets rid of freeze
    // sheet.createFreezePane( 0, 1 );  // first number col, 2nd row
  }

  /**
   * Set the sheet margins to fairly narrow
   */
  public void setSheetMargins() {
    if (sheet == null) { return; }
    sheet.setMargin((short)0, .5d);
    sheet.setMargin((short)1, .5d);
    sheet.setMargin((short)2, .5d);
    sheet.setMargin((short)3, .5d);
  }

  /**
   * To be called after the sheet is filled in.  Set all the columns to their
   * preferred width based on content.
   */
  public void setSheetColumnWidths() {
    if (sheet == null) { return; }
    Row rowX = sheet.getRow(0);
    if (rowX == null) { return; }
    for (int i = 0; i<rowX.getLastCellNum(); i++) {
      sheet.autoSizeColumn(i);
    }
  }

  /**
   * The only row in the sheet is the column headings.  Set the column
   * widths to accommodate the column labels.
   */
  public void setColumnWidthsToLabels() {
    if ((sheet == null) || (row == null)) { return; }
    /* The last cell number is one more than the index of the last
     * existing cell*/
    for (int k = row.getFirstCellNum(); k<row.getLastCellNum(); k++) {
      sheet.autoSizeColumn(k);
    }
  }

  /**
   * Process all the rows based on having explored the sheet characteristics
   * @return the next non-comment row in the current sheet or null if
   * there are no more rows.
   */
  public Row nextRow() {
    if (currentRowNumber >= lastRowNumber) { return null; }
    Cell cell;
    do {
      currentRowNumber++;
      row = sheet.getRow(currentRowNumber);
      if (row != null) {
	cell = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
	if (cell == null) { continue; }
	if ((cell.getCellType() == CellType.NUMERIC) ||
	    ((cell.getStringCellValue() != null) &&
		!cell.getStringCellValue().startsWith("#"))) { return row; }
      }
    } while (currentRowNumber <= lastRowNumber);
    return null;
  }

  /**
   * @return a new row which was added to the work sheet.
   */
  public Row addRowToSheet() {
    return row = sheet.createRow(sheet.getLastRowNum()+1);
  }

  /**
   * Create a column map from the current row, ignoring empty columns.
   * @return list of duplicate column names or null if there are none.
   */
  public List<String> duplicateColumns() {
    List<String> dupColNames = null;
    String content;
    Cell cell;
    int cellCount;
    int cellIdx;

    columnMap = new HashMap<String, Integer>();

    cellCount = row.getLastCellNum();
    for (cellIdx=0; cellIdx<cellCount; cellIdx++) {
      if ( (cell = row.getCell(cellIdx)) == null) { continue; }
      content  = SSU.getCellValue(cell, false).trim();
      if ((content != null) && (content.length() > 0)) {
	if (columnMap.get(content) != null) {
	  if (dupColNames == null) { dupColNames = new ArrayList<String>(); }
	  dupColNames.add(content);
	} else {
	  columnMap.put(content, new Integer(cellIdx));
	}
      }
    }
    return dupColNames;
  }

  /**
   * @param cols array of required column names.
   * @return list of missing columns or null if there are no missing
   * columns
   */
  public List<String> missingColumns(String[] cols) {
    List<String> missing = null;
    for (String col : cols) {
      if (!columnMap.containsKey(col)) {
	if (missing == null) {
	  missing = new ArrayList<String>();
	}
	missing.add(col);
      }
    }
    return missing;
  }

  /**
   * Make sure that the current work sheet has at least 2 rows and
   * that the first physical row has a list of column names which have
   * no duplicates and which contain all the required column names
   * which are listed in the input column name array
   * @param requiredColNames required columns
   * @return true if the work sheet is OK, false otherwise.
   */
  public boolean isSheetOK(String[] requiredColNames) {
    return isSheetOK(requiredColNames, 2);
  }

  /**
   * Make sure that the current work sheet has enough rows and
   * that the first physical row has a list of column names which have
   * no duplicates and which contain all the required column names
   * which are listed in the input column name array
   * @param requiredColNames required columns
   * @param minRows Minimum number of rows in the sheet
   * @return true if the work sheet is OK, false otherwise.
   */
  public boolean isSheetOK(String[] requiredColNames, int minRows) {
    List<String> badCol;

    if ((row == null) || (physicalRows < minRows)) {
      sb.append("Work sheet " + sheetName +
	  (isStringMTP(fileName) ? "" : " in " + fileName) + " must have " + minRows +
	  " or more rows<br>\n");
      return false;
    }
    badCol = duplicateColumns(); // Creates column map
    if (badCol != null) {
      augmentComplaints("Work sheet " + sheetName +
	  (isStringMTP(fileName) ? "" : " in " + fileName) +
	  " has duplicate column names",
	  badCol);
      return false;
    }
    badCol = missingColumns(requiredColNames);
    if (badCol != null) {
      augmentComplaints("Work sheet " + sheetName +
	  (isStringMTP(fileName) ? "" : " in " + fileName) +
	  " is missing these columns",
	  badCol);
      return false;
    }
    return true;
  }

  /**
   * See if all required values in the current row have values of some
   * sort
   * @param colsNeedingValues list of columns which must have values
   * @return true if all columns have values, false otherwise
   */
  public boolean isRowOK(String[] colsNeedingValues) {
    List<String> badCols = missingValues(colsNeedingValues);
    if (badCols != null) {
      augmentComplaints("These columns in worksheet " + sheetName + " row " +
	  (currentRowNumber+1) +
	  " must have values", badCols);
      return false;
    }
    return true;
  }

  /**
   * @param cols list of column names
   * @return null if no columns have missing values or a list of
   * columns whose values are missing in the current row.  A null cell
   * or a non-null value which has no characters is empty.
   */
  public List<String> missingValues(String[] cols) {
    List<String> missing = null;
    Integer which;
    Cell cell;
    String value;
    for (String col : cols) {
      if ( (which = columnMap.get(col)) != null) {
	cell = row.getCell(which);
	if (cell != null) {
	  value = SSU.getCellValue(cell, false);
	  if ((value != null) && (value.length() > 0)) {
	    continue;
	  }
	}
      }
      if (missing == null) {
	missing = new ArrayList<String>();
      }
      missing.add(col);
    }
    return missing;
  }

  /**
   * Augment an error message string
   * @param why why the list of bad strings is being created
   * @param what list of bad strings
   */
  public void augmentComplaints(String why, List<String> what) {
    sb.append(why + ":");
    for (String w : what) {
      sb.append("\t" + w);
    }
    sb.append("<br>\n");
  }

  /**
   * Add a complaint to the accumulated string
   * @param whinge text of the complaint
   */
  public void whinge(String whinge) {
    whinge(whinge, true);
  }

  /**
   * Add a complaint to the accumulated string
   * @param whinge text of the complaint
   * @param wantRow add sheet name and row number to whinge
   */
  public void whinge(String whinge, boolean wantRow) {
    sb.append((wantRow ? "Sheet " + sheetName + " row " + (currentRowNumber + 1) + " " : "") +
	whinge + "<br>\n");
  }

  /**
   * @param is input string
   * @return true if the string is either null or empty
   */
  public boolean isStringMTP(String is) {
    return (is == null) || (is.length() <= 0);
  }

  /**
   * @param col name of the column
   * @return column value formatted as Excel would format it or null
   * if there is an error trying to get a nonexistent column or other
   * excel issue.
   */
  public String getFormattedColumn(String col) {
    try {
      return SSU.getFormattedCell(columnMap.get(col), row);
    } catch (Exception e) {
      return null;
    }
  }

  /** Check a boolean column for matching a boolean value.  If they are
   * not the same, set the desired value and set the dirty bit
   * @param col column name.
   * @param what what the value should be.
   * @param dirtyIn add to error messages if true
   */
  public void setBoolAndDirty(String col, boolean what, boolean dirtyIn){
    String want = what ? "TRUE" : "false";
    String got = getFormattedColumn(col);
    if (want.equalsIgnoreCase(got)) { return; }
    dirty = true;
    setMappedCellStringValue("Status", "X");
    setMappedCellStringValue(col, want, isStringMTP(got) ? null : headerTextStyle);
    if (dirtyIn) { whinge(col); }
  }

  /** Set a string value if the cell content is not the same as the
   * wanted string and set the dirty bit
   * @param col column name
   * @param want desired string value
   * @param dirtyIn complain if the cell value changes
   */
  public void setStringAndDirty(String col, String want, boolean dirtyIn) {
    String got = getFormattedColumn(col);
    if ((got != null) && got.equalsIgnoreCase(want)) { return; }
    dirty = true;
    setMappedCellStringValue("Status", "X");
    setMappedCellStringValue(col, want, isStringMTP(got) ? null : headerTextStyle);
    if (dirtyIn) { whinge(col); }
  }

  /** Conditionally set an integer value in a column.  Does nothing if the
   * column already has the desired value.
   * @param col column to get a new value
   * @param want desired value for the column
   * @param dirtyIn adds the column name to the list of complaints if
   * column value is different and this flag is true.
   */
  public void setIntegerAndDirty(String col, int want, boolean dirtyIn) {
    int got = getIntegerColumn(col);
    if (got == want) { return; }
    dirty = true;
    setMappedCellStringValue("Status", "X");
    Cell cell = row.getCell(columnMap.get(col), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
    cell.setCellType(CellType.NUMERIC);
    cell.setCellStyle(integerSty);
    cell.setCellValue(want);
    if (dirtyIn) { whinge(col); }
  }

  /** Return the value of a column treating it as an integer
   * @param col column name
   * @return integer value or 0 if it is not an integer
   */
  public int getIntegerColumn(String col) {
    try {
      return Integer.valueOf(getFormattedColumn(col));
    } catch (NumberFormatException e) { /* */ }
    return 0;
  }

  /** Retrieve the value of a cell as an excel formatted date
   * @param col column name
   * @return formatted date value
   */
  public String getDateColumn(String col) {
    return SSU.df.formatCellValue(row.getCell(columnMap.get(col)));
  }

  /**
   * Set the value of the named cell in the current row.  Does nothing if
   * that cell is null.  This is intended for overriding existing cell
   * values.
   * @param col name of the column
   * @param val value to set in the cell
   */
  public void setColumnTo(String col, String val) {
    int which = columnMap.get(col);
    Cell cell = row.getCell(which, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
    if (cell != null) {
      cell.setCellValue(val);
    }
  }

  /**
   * Give a named column a string value
   * @param col column name
   * @param value string value for the cell
   */
  public void setMappedCellStringValue(String col, String value) {
    setMappedCellStringValue(col, value, null);
  }

  /**
   * Give a named column a string value
   * @param col column name
   * @param value string value for the cell
   * @param cellStyle determines the cell style if not null
   */
  public void setMappedCellStringValue(String col, String value, CellStyle cellStyle) {
    int which = 0;
    try {
      which = columnMap.get(col);
    } catch (Exception e) {
      whinge("Missing column " + col, true);
      return;
    }
    Cell cell = row.getCell(which, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
    cell.setCellValue(value);
    cell.setCellType(CellType.STRING);
    if (cellStyle != null) {
      cell.setCellStyle(cellStyle);
    }
  }

  @Override
  public String toString() {
    return "Sheet " + sheetName + " fr " + firstRowNumber +
	" lr " + lastRowNumber;
  }
}
