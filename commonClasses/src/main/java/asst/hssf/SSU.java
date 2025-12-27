package asst.hssf;

import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import asst.dbcommon.PUTs;

/**
 * Spread Sheet Utilities for processing either .xls or .xlxs
 * files
 * @author Material Gain
 * @since 2015 01
 */
public class SSU {

  static final long serialVersionUID = 1;

  /**
   * Format a date object as an Excel-compliant date
   */
  public static final SimpleDateFormat EXCEL_DATE =
      new SimpleDateFormat("MM/dd/yyyy");
  /** Format for writing to a database*/
  public static final SimpleDateFormat SQL_DATE_OUTPUT =
      new SimpleDateFormat("yyyy-MM-dd");
  /** Format for writing date and time to database*/
  public static final SimpleDateFormat SQL_DATE_TIME_OUTPUT =
      new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.S");
  /** Turn any valid number string into a Number whose value is an integer.
   * The output cannot be cast to either an int or to a long, it has to be
   * a Number, but that can be processed in various ways.  */
  public static final NumberFormat INT_ONLY =
      NumberFormat.getIntegerInstance();

  /**
   * Public data formatter for use with POI-aware programs
   */
  public static final DataFormatter df = new DataFormatter();

  /** Return a work sheet with a specified name.  Creates the work
   * sheet if it does not exist.
   * @param name name of the work sheet
   * @param freezeH true means to freeze the top row
   * @param columns array of column names, if null, no title
   * row is generated
   * @param outBook workbook where the sheet is created
   * @return work sheet
   */
  public static Sheet getOrMakeSheet(String name, boolean freezeH,
      String[] columns, Workbook outBook) {
    Sheet outSheet = outBook.getSheet(name);
    if (outSheet == null) {
      outSheet = outBook.createSheet(name);
      SSU.printlnSheet(outSheet, columns, true, true);
      /* sheet.createFreezePane( 0, 1, 0, 1 ); for a row,
       * sheet.createFreezePane( 1, 0, 1, 0 ); for a column */
      if (freezeH) {
	outSheet.createFreezePane( 0, 1, 0, 1 );
      }
    }
    return outSheet;
  }

  /**
   * Remove a named work sheet from a spread sheet
   * @param sName the name of the sheet to get rid of
   * @param outBook the work book which may contain the sheet
   */
  public static void removeSheet(String sName, Workbook outBook) {
    Sheet sheet = outBook.getSheet(sName);
    if (sheet == null) { return; }
    int sheetIx = outBook.getSheetIndex(sheet);
    outBook.removeSheetAt(sheetIx);
  }

  /**
   * Give a named column a string value
   * @param col column name
   * @param row current row
   * @param value string value for the cell
   * @param colMap Column map
   */
  public static void setMappedCellStringValue(String col, Row row,
      String value, Map<String, Integer> colMap) {
    Cell cell;
    Integer colNo;
    if ( (colNo = colMap.get(col)) == null) {
      throw new RuntimeException("Column " + col +
	  " is not in the column map.<br>");
    }
    cell = row.getCell(colNo, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
    cell.setCellValue(value);
  }

  /**
   * Return a string representing the cell value
   * @param cell spread sheet cell
   * @param isDate true means to expect a date in the cell
   * @return a blank string if the cell is null or cannot be formatted
   * as a date and is flagged as a date
   */
  public static String getCellValue(Cell cell, boolean isDate) {
    if (cell == null) { return ""; }
    Date date = null;

    /* Unfortunately, the POI utilities do not have a value to
     * indicate a date type column.  Empty cells return null when
     * attempts are made to convert them to dates; non-date columns
     * cause exceptions.  It may be necessary to KNOW which cells
     * should be dates and handle them accordingly.*/
    if (isDate) {
      try {
	if ( (date = cell.getDateCellValue()) == null) { return ""; }
	return EXCEL_DATE.format(date);
      } catch (Exception e) {
	// If not a date, just proceed to normal processing
      }
    }

    switch (cell.getCellType()) {

    case NUMERIC:
      return (String.valueOf(cell.getNumericCellValue()));

    case STRING:
      return(cell.getRichStringCellValue().getString().trim());

    case ERROR:
      return "";

    case BLANK:
      return "";

    case FORMULA:
      String s4 = cell.getCellFormula();
      /* A hyper link formula takes 2 args, the second is the display
       * value.  These args are enclosed in quotes.*/
      if (s4.startsWith("HYPERLINK")) {
	int ix = s4.indexOf(",");
	if (ix > 0) {
	  return s4.substring(ix+2, s4.length()-2);
	}
      }
      return s4;
    case BOOLEAN:
      break;
    case _NONE:
      return "";
    default:
      return "";
    }

    return "";
  }

  /** Return a column number based on the column name or -1 if the
   * column is not found
   * @param col name of the column
   * @param colMap map column names to column numbers
   * @return number of the column or -1 if it is not found or there is
   * no map
   */
  public static int getColumnNumber(String col, Map<String, Integer> colMap) {
    Integer idx;
    if ((colMap == null) || ( (idx = colMap.get(col)) == null)) {
      return -1;
    }
    return idx.intValue();
  }

  /**
   * Return the value of the cell based on Excel's internal formatting
   * rules.  Cell 0 is treated in a special manner - if it starts with
   * #, the row is treated as a comment row
   * @param column column number which may be out of range of the row.
   * Column zero is treated in a special manner in that it may
   * indicate a comment row.
   * @param row which may be null in which case, returns null for any cell
   * @return cell value formatted as excel would display it.  In
   * particular, this preserves integers as integers.  Returns null if
   * column==0 and the string starts with # which indicates a comment
   * row.  An empty cell returns the empty string, a null cell returns
   * null.
   */
  public static String getFormattedCell(int column, Row row) {
    if (row == null) { return null; }
    Cell cell = row.getCell(column);
    if (cell == null) { return null; }
    String content = df.formatCellValue(cell);
    if ((column == 0) && content.startsWith("#")) { return null; }
    return content;
  }

  /**
   * return the column as Excel formats it
   * @param column name
   * @param row which should have an instance of the column
   * @param colMap maps column names to numbers
   * @return column content
   */
  public static String getFormattedColumn(String column, Row row,
      	Map<String, Integer> colMap) {
    return getFormattedCell(getColumnNumber(column, colMap), row);
  }

  /**
   * Test a column value to be sure that Excel has formatted it as a valid date
   * @param column name
   * @param row which has the column
   * @param inFuture true means that the date must not be in the past
   * @param colMap maps column names to numbers
   * @return null if the column value is OK, otherwise returns the bad value.
   */
  public static String isValidDateP(String column, Row row,
    	boolean inFuture, Map<String, Integer> colMap) {
    String val = getFormattedColumn(column, row, colMap);
    java.util.Date date;
    try {
      date = EXCEL_DATE.parse(val);
      if (inFuture) {
	if (date.getTime() < System.currentTimeMillis()) {
	  return val + " is in the past";
	}
      }
      return null; // the date is OK
    } catch (Exception e) {}
    return val; // The date string is bad
  }

  /** Convert an Excel date to a string suitable for writing to a database.
   * @param xlDat date in excel format
   * @return date string in database format.*/
  public static String dbDateFroExcel(String xlDat) {
    java.util.Date date;
    try {
      date = EXCEL_DATE.parse(xlDat);
      return SQL_DATE_OUTPUT.format(date);
    } catch (Exception e) {
      date = new java.util.Date();
    }
    return SQL_DATE_OUTPUT.format(date);
  }

  /** Return the value of a cell as a string based on its cell number
   * in the row.  Cell 0 is treated in a special manner - if it starts
   * with #, the row is treated as a comment row
   * @param column column number, which may be out of range of the
   * row.  Column zero is treated in a special manner in that it may
   * indicate a comment row.
   * @param row which may be null
   * @return cell content if it is a non-empty string otherwise null
   * if the cell is null, its content is null or the empty string, or
   * if column==0 and the string starts with # which indicates a
   * comment row*/
  public static String getNumberedCell(int column, Row row) {
    if (row == null) { return null; }
    Cell cell = row.getCell(column);
    if (cell == null) { return null; }
    String content = df.formatCellValue(cell);
    if (content.length() <= 0) { return null; }
    if ((column == 0) && content.startsWith("#")) { return null; }
    return content.trim();
  }

  /** Return the value of a cell as a string which represents an
   * integer with any trailing data past a decimal point dropped.
   * This is because poi returns integer values with at least one
   * decimal past the period and sometimes with commas.
   * @param column number, which may be out of range of the row.
   * Column zero is treated in a special manner in that it may
   * indicate a comment row in which case a value of 0 is returned.
   * @param row which may be null
   * @return string as an integer with decimal point chopped off*/
  public static String getIntFromNumberedCell(int column, Row row) {
    String val = getNumberedCell(column, row);
    return makeInt(val);
  }

  /**
   * POI tends to return integers as strings of the form xx.0.  These
   * strings do nor parse properly as integers.
   * @param in data from an excel cell
   * @return integer value of the string, or 0 if input string is null
   * or empty.
   */
  public static String makeInt(String in) {
    if (PUTs.isStringMTP(in)) { return "0"; }
    try {
      Number value = INT_ONLY.parse(in);
      in = String.valueOf(value);
    } catch (ParseException e) {
      return "0";
    }
    return in;
  }

  /** Set the font in all cells to TimesNewRoman, then adjust all the
   * cell sizes in all sheets of the work book so that the column
   * widths can display the cell content.
   * @param book the work book all of whose work sheets are to be
   * sized so that all columns are of the default width to show all their
   * content.
   */
  public static void autoSizeAll(Workbook book) {
    Sheet sheet;
    Row row;
    Cell cell;
    CellStyle style = null;
    int i;
    int j;
    int k;
    int sheets = book.getNumberOfSheets();
    int rows;
    int cols;

    for (i=0; i<sheets; i++) {
      sheet = book.getSheetAt(i);
      rows  = sheet.getLastRowNum();
      for (j=0; j<=rows; j++) {
	row = sheet.getRow(j);
	if (row == null) { continue; }
	cols = row.getLastCellNum();
	for (k=0; k<=cols; k++) {
	  if ( (cell = row.getCell(k)) == null) { continue; }
	  if (style == null) { // Take the style from the first non-null cell
	    style = cell.getCellStyle();
	    Font font = book.getFontAt(style.getFontIndex());
	    font.setFontName("TimesNewRoman");
	  } else {
	    //cell.setCellStyle(style); // Not needed - cells share the same style
	  }
	}
      }
    }

    for (i=0; i<sheets; i++) {
      sheet = book.getSheetAt(i);
      sheet.setMargin((short)0, .5d);
      sheet.setMargin((short)1, .5d);
      sheet.setMargin((short)2, .5d);
      sheet.setMargin((short)3, .5d);
      k = 0;
      rows  = sheet.getLastRowNum();
      for (j=0; j<=rows; j++) {
	row = sheet.getRow(j);
	if (row == null) { continue; }
	cols = row.getLastCellNum();
	for (; k<=cols; k++) {
	  sheet.autoSizeColumn((short) k);
	}
      }
    }
  }

  /**
   * Copy the contents of a row of a spread sheet to the next row of a
   * different spread sheet.  The input row may or may not be in the
   * spread sheet to which it is being copied
   * @param outSheet spread sheet work sheet which is to contain a
   * copy of the input row
   * @param inRow input row which may be null
   * @return output row which will be empty if the input row is null,
   * otherwise contains cells whose text values are the same as the
   * text values extracted from successive cells of the input row.
   * Null cells in the input row are skipped and remain null in the
   * output row.
   */
  public static Row copyRowToSheet(Sheet outSheet, Row inRow) {
    Row outputRow = outSheet.createRow(outSheet.getPhysicalNumberOfRows());
    Cell cell;
    int i;

    if (inRow != null) {
      for (i=0; i<=inRow.getLastCellNum(); i++) {
	if ( (cell = inRow.getCell(i)) == null) { continue; }
	outputRow.createCell(i).setCellValue(getCellValue(cell, false));
      }
    }
    return outputRow;
  }

  /**
   * Treat a work sheet as if it were an output stream to be printed
   * in successive rows by printing after or on the last existing row
   * @param sheet the work sheet
   * @param pl the array of strings to be put in the next row, one per cell.
   * The array must be at least of length one.
   * @return the new spread sheet row with the cells filled with the
   * strings
   */
  public static Row printlnSheet(Sheet sheet, String[] pl) {
    return printlnSheet(sheet, pl, false, false);
  }

  /**
   * Treat a work sheet as if it were an output stream to be printed
   * in successive rows by printing after or on the last existing row
   * @param sheet the work sheet
   * @param pl the array of strings to be put in the next row, one per cell.
   * The array must be at least of length one.
   * @param noUnder if true, all underscores are converted to spaces
   * @param on  true means to print on the last row instead of after it
   * @return the new spread sheet row with the cells filled with the
   * strings
   */
  public static Row printlnSheet(Sheet sheet, String[] pl, boolean noUnder,
      boolean on) {
    if (pl == null) { return null; }
    String c = pl[0];
    if ((c != null) && noUnder) { c = c.replace('_', ' '); }
    Row row = printlnSheet(sheet, c, on);
    for (int i=1; i<pl.length; i++) {
      c = pl[i];
      if ((c != null) && noUnder) { c = c.replace('_', ' '); }
      row.createCell(i).setCellValue(c);
    }
    return row;
  }

  /**
   * Treat a work sheet as if it were an output stream to be printed
   * in successive rows by printing after or on the last existing row
   * @param sheet the work sheet
   * @param pl the string to be put in the next row
   * @param on true means to print on the last row instead of after it
   * @return the new spread sheet row with the first cell filled with the
   * message string
   */
  public static Row printlnSheet(Sheet sheet, String pl, boolean on) {
    Row row = sheet.createRow(sheet.getLastRowNum()+(on ? 0 :1));
    row.createCell(0).setCellValue(pl);
    return row;
  }
}
