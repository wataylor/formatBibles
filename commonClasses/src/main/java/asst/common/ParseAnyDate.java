package asst.common;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.TimeZone;

/**
 * Parse any of many different date formats, remembering the most recent
 * successful parse string to try first on the next attempt.
 * @author Material gain
 * @since 2017 02
 *
 */
public class ParseAnyDate {

  TimeZone tz;
  //Remember the most recent successful parse string
  SimpleDateFormat recent = null;

  private static final Map<String, String> DATE_FORMAT_REGEXPS =
      new HashMap<String, String>() {
    private static final long serialVersionUID = 1L;

    {
      put("^\\d{8}$", "yyyyMMdd");
      put("^\\d{1,2}-\\d{1,2}-\\d{2}$", "MM-dd-yy");
      put("^\\d{1,2}/\\d{1,2}/\\d{2}$", "MM/dd/yy");
      put("^\\d{1,2}-\\d{1,2}-\\d{4}$", "MM-dd-yyyy");
      put("^\\d{4}-\\d{1,2}-\\d{1,2}$", "yyyy-MM-dd");
      put("^\\d{1,2}/\\d{1,2}/\\d{4}$", "MM/dd/yyyy");
      put("^\\d{4}/\\d{1,2}/\\d{1,2}$", "yyyy/MM/dd");
      put("^\\d{1,2}\\s[a-z]{3}\\s\\d{4}$", "dd MMM yyyy");
      put("^\\d{1,2}\\s[a-z]{4,}\\s\\d{4}$", "dd MMMM yyyy");
      put("^\\d{12}$", "yyyyMMddHHmm");
      put("^\\d{8}\\s\\d{4}$", "yyyyMMdd HHmm");
      put("^\\d{1,2}-\\d{1,2}-\\d{4}\\s\\d{1,2}:\\d{2}$", "dd-MM-yyyy HH:mm");
      put("^\\d{4}-\\d{1,2}-\\d{1,2}\\s\\d{1,2}:\\d{2}$", "yyyy-MM-dd HH:mm");
      put("^\\d{1,2}/\\d{1,2}/\\d{4}\\s\\d{1,2}:\\d{2}$", "MM/dd/yyyy HH:mm");
      put("^\\d{4}/\\d{1,2}/\\d{1,2}\\s\\d{1,2}:\\d{2}$", "yyyy/MM/dd HH:mm");
      put("^\\d{1,2}\\s[a-z]{3}\\s\\d{4}\\s\\d{1,2}:\\d{2}$", "dd MMM yyyy HH:mm");
      put("^\\d{1,2}\\s[a-z]{4,}\\s\\d{4}\\s\\d{1,2}:\\d{2}$", "dd MMMM yyyy HH:mm");
      put("^\\d{14}$", "yyyyMMddHHmmss");
      put("^\\d{8}\\s\\d{6}$", "yyyyMMdd HHmmss");
      put("^\\d{1,2}-\\d{1,2}-\\d{4}\\s\\d{1,2}:\\d{2}:\\d{2}$", "dd-MM-yyyy HH:mm:ss");
      put("^\\d{4}-\\d{1,2}-\\d{1,2}\\s\\d{1,2}:\\d{2}:\\d{2}$", "yyyy-MM-dd HH:mm:ss");
      put("^\\d{1,2}/\\d{1,2}/\\d{4}\\s\\d{1,2}:\\d{2}:\\d{2}$", "MM/dd/yyyy HH:mm:ss");
      put("^\\d{4}/\\d{1,2}/\\d{1,2}\\s\\d{1,2}:\\d{2}:\\d{2}$", "yyyy/MM/dd HH:mm:ss");
      put("^\\d{1,2}\\s[a-z]{3}\\s\\d{4}\\s\\d{1,2}:\\d{2}:\\d{2}$", "dd MMM yyyy HH:mm:ss");
      put("^\\d{1,2}\\s[a-z]{4,}\\s\\d{4}\\s\\d{1,2}:\\d{2}:\\d{2}$", "dd MMMM yyyy HH:mm:ss");
    }};

    /** Default constructor */
    public ParseAnyDate() {
      
    }

    /**
     * @param tz desired time zone for date parsing.  Uses
     * the default time zone for the locale if this is null.
     */
    public ParseAnyDate(TimeZone tz) {
      this.tz = tz;
    }

    /**
     * Determine SimpleDateFormat pattern matching with the given date string. 
     * Extend DateUtil with more formats if needed.
     * @param dateString The date string which needs a SimpleDateFormat pattern.
     * @return Matching SimpleDateFormat pattern or null if format is unknown.
     * @see java.text.SimpleDateFormat
     */
    public static String determineDateFormat(String dateString) {
      for (String regexp : DATE_FORMAT_REGEXPS.keySet()) {
	if (dateString.toLowerCase().matches(regexp)) {
	  return DATE_FORMAT_REGEXPS.get(regexp);
	}
      }
      return null; // Unknown format.
    }

    /**
     * Find a date parser for the input date string
     * @param dateString  date in one of many supported notations
     * @return Date object or null
     */
    public SimpleDateFormat getDateFormatter(String dateString) {
      if (recent != null) {
	return recent;
      }
      String formatString = determineDateFormat(dateString);
      if (formatString == null) {
	return null;
      }
      recent = new SimpleDateFormat(formatString);
      if (tz != null) {
	recent.setTimeZone(tz);
      }
      return recent;
    }

    /**
     * Try very hard to parse an input date
     * @param dateString date in one of many supported notations
     * @return Date object or null
     */
    public Date parseAnyDate(String dateString) {
      try {
	return recent.parse(dateString);
      } catch (Exception e) {
	recent = null;
	try {
	  return getDateFormatter(dateString).parse(dateString);
	} catch (Exception e1) {
	}
      }
      return null;
    }
}
