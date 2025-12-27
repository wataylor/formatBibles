package asst.hssf;

import java.text.SimpleDateFormat;

/**
 * Hold the result of analyzing a date.  Spread Sheet dates must be
 * mm/dd/yyyy precisely.  January must be 01, not 1 for example.
 * @author Material Gain
 * @since 2105 02
 */
public class DateResult {

  /**
   * Format a date object as an Excel-compliant date
   */
  public static final SimpleDateFormat EXCEL_DATE =
      new SimpleDateFormat("MM/dd/yyyy");

  /** This is the original date string which was input by the
   * caller */
  public String originalDate;
  /** This is the properly-formatted date regardless of errors about
   * the future. */
  public String usableDate;
  /** The error message string is non-null if there is something wrong
   * with the date. */
  public String msg;
  boolean inFuture;

  /**
   * Return a date object based on its formatting.
   * @param originalDate a date string which should be formatted
   * properly but might not be.
   */
  public DateResult(String originalDate) {
    this.originalDate = originalDate;
    analyze();
  }

  /**
   * Return a date object based on its formatting.
   * @param originalDate a date string which should be formatted
   * properly but might not be.
   * @param inFuture true means that the date may not be in the past.
   * If false, past or future does not matter.
   */
  public DateResult(String originalDate, boolean inFuture) {
    this(originalDate);
    this.inFuture = inFuture;
    analyze();
  }

  private void analyze() {
    java.util.Date date;
    try {
      date = EXCEL_DATE.parse(originalDate);
      if (inFuture) {
	if (date.getTime() < System.currentTimeMillis()) {
	  msg = originalDate + " is in the past";
	}
      }
      if (date.getYear() < 110) { // returns years after 1900
	msg = originalDate + " is not formatted mm/dd/yyyy";
      }
      usableDate = EXCEL_DATE.format(date);
    } catch (Exception e) {
      msg = originalDate + " is not formatted mm/dd/yyyy";
    }
  }
}
