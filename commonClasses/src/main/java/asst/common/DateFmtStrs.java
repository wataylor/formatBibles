package asst.common;

import java.text.NumberFormat;
import java.util.TimeZone;

/** Define useful date formatting strings so that threads can easily create
 * their own date formatters to bypass threading issues in
 * the SimpleDateFormatter class.  When used in single-thread batch
 * programs, thread safety is not an issue.
 * D 	Day in year
 * d 	Day in month
 * H 	Hour in day (0-23)
 * h 	Hour in am/pm (1-12)
 * a 	Am/pm marker
 * u 	Day number of week (1 = Monday, ..., 7 = Sunday)
 * @author Material Gain
 * @since 2019 09
 */
public class DateFmtStrs {
  /** Used to format times when users clicked or redeemed coupons. */
  public static final String USER_TIME_AT_FORMATTER = "E, MMMM d, yyyy 'at' h:mm a";
  /** Used to format times when next redemption starts. */
  public static final String USER_TIME_AROUND_FORMATTER = "E, MMMM d, yyyy 'around' h:mm a";
  /** Used to format times when next redemption starts. */
  public static final String USER_TIME_AROUND_MILITARY = "E, MMMM d, yyyy 'around' HH:mm a";
  /** Used to format times when a coupon expires. */
  public static final String MONTH_NAME_DAY_YEAR = "MMMM d, yyyy";
  /** Used to format times when experimenting with day of week. */
  public static final String MONTH_NAME_DAY_NAME_YEAR = "EEE MMMM d, yyyy";
  /** Used to format user-entered click times. */
  public static final String yyyy_MM_d_hh_mm = "yyyy-MM-d hh:mm";
  /** Format in which coupon expiration date strings are stored in the database*/
  public static final String MM_dd_yyyy = "MM/dd/yyyy";
  /** Just the hour and minute*/
  public static final String HH_MM_a = "h:mm a";
  /** Sortable date / time with hours, minutes, seconds, and time zone. */
  public static final String SORTABLE_WITH_SECONDS_AND_TZ = "yyyy-MM-dd HH:mm:ss Z";
  /** Sortable date / time with hours, minutes, seconds, time zone, and day. */
  public static final String SORTABLE_WITH_SECONDS_TZ_DAY = "yyyy-MM-dd HH:mm:ss Z EEE";
  /** Sortable date / time with hours, minutes, and seconds. */
  public static final String SORTABLE_WITH_SECONDS = "yyyy-MM-dd HH:mm:ss";
  /** Short sortable */
  public static final String yyyy_MM_dd = "yyyy-MM-dd";
  /** Another way to describe this string for reading Excel dates */
  public static final String EXCEL_DATE_STRING = "MM/dd/yyyy";
  /** String for Excel date time cells */
  public static final String EXCEL_DATE_TIME_STRING = "MM/dd/yyyy hh:mm a";
  /** Another string for Excel date time cells */
  public static final String EXCEL_DATE_TIME_STRING_SECONDS = "MM/dd/yyyy hh:mm:ss a";
  /** Produce and parse date from form selector boxes to write to MySQL.  */
  public static final String SQL_DATE_TIME_OUTPUT_STRING = "yyyy-MM-dd HH:mm:ss";
  /** An acceptable way to format a date for writing into MySQL.  */
  public static final String SQL_DATE_OUTPUT_STRING = "yyyy-MM-dd";

  /** Format a single-digit number with a leading zero.*/
  public static NumberFormat LZERO_NUMBER = NumberFormat.getInstance();
  /** Format a single-digit number with up to 3 leading zeros.*/
  public static NumberFormat L3ZERO_NUMBER = NumberFormat.getInstance();
  static {
    LZERO_NUMBER.setMinimumIntegerDigits(2);
    L3ZERO_NUMBER.setMinimumIntegerDigits(3);
  }

  /** GMT time zone.  */
  public static final TimeZone zuluTZ = TimeZone.getTimeZone("zulu");

}
