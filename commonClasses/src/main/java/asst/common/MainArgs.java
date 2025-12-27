/* @name MainArgs.java

   Copyright (c) 2002-2011 Advanced Systems and Software Technologies, LLC
   (All Rights Reserved)

-------- Licensed Software Proprietary Information Notice -------------

This software is a working embodiment of certain trade secrets of
AS&ST.  The software is licensed only for the day-to-day business use
of the licensee.  Use of this software for reverse engineering,
decompilation, use as a guide for the design of a competitive product,
or any other use not for day-to-day business use is strictly
prohibited.

All screens and their formats, color combinations, layouts, and
organization are proprietary to and copyrighted by AS&ST, LLC.

All rights are reserved.

Authorized AS&ST customer use of this software is subject to the terms
and conditions of the software license executed between Customer and
AS&ST, LLC.

------------------------------------------------------------------------

*/

package asst.common;

import java.util.Hashtable;
import java.util.LinkedList;
import java.util.List;

/** Support Command Line argument parsing for Java main programs.
 *
 * @author Material Gain
 * @since 2005 01
 */

public class MainArgs extends Hashtable<String, Object> {

  static final long serialVersionUID = 1;

  List<String> strings=new LinkedList<String>();

  /** Obligatory constructor.*/
  public MainArgs() { /* */ }

  /**
 * @param args a copy of the command line arguments or a string array
 * to set defaults.
 */
public MainArgs(String[] args) {
    parseArgs(args);
  }

  /** Add whatever arguments are specified by the string array to the
   * hash table.  The constructor can call this to set defaults,
   * the arguments can be passed to this method to override defaults if any
   * @param args array of argument strings.  Quotes are removed and
   * null or zero-length entries are skipped.  */
  public void parseArgs(String[] args) {
    String param;
    for(String arg : args) {
      if (arg == null) { continue; }
      // If a string begins with a quote, assume it has two and remove both
      if (arg.startsWith("\"")) { arg = arg.substring(1, arg.length()-1); }

      // Skip empty strings
      if (arg.length() <= 0) { continue; }

      param=arg.substring(1);
      if(arg.indexOf('=')>=0) {
        put(arg.substring(0, arg.indexOf('=')),
	    arg.substring(arg.indexOf('=')+1));
      } else if(arg.charAt(0)=='-') {
        put(param, Boolean.FALSE);
      } else if(arg.charAt(0)=='+') {
        put(param, Boolean.TRUE);
      } else {
        strings.add(arg);
      }
    }
  }

  /** Parameters which are not of the form name=value are stored in a
   * list in the order they are encountered on the command line.  This
   * method returns a list of all such command line parameters.
   * @return list of strings which were defined without a keyword*/
  public List<String> getStrings() {
    return strings;
  }

  /** Return the value of a parameter as a boolean; defaults to
   * <code>false</code> unless the parameter has been entered as
   * +&lt;param name&gt; in which case, it returns true
   * @param param name of the parameter to be returned.
   * @return boolean value of the keyword.*/
  public boolean getBoolean(String param) {
    Boolean b;
    if ( (b = (Boolean)get(param)) == null) { return false; }
    return b.booleanValue();
  }

  /**
   * Convert a string to an integer, returning 0 on failure
   * @param va string which represents an integer
   * @return the value of the integer or 0 if the string is null or
   * does not represent a properly-formatted integer.
   */
  public static int integerFromString(String va) {
    try {
      return Integer.decode(va);
    } catch (Exception e) { /* */ }
    return 0;
  }

  /**
   * Get the value of a parameter as an integer.  Returns 0 if the
   * parameter is not found or if it is not a properly-formatted
   * integer
   * @param param name of the parameter to be converted to an integer
   * @return value of the integer or 0 if the parameter is not found
   */
  public int getInt(String param) {
    return integerFromString((String)get(param));
  }

  /**
   * Convert a string to a double, returning 0 if the string is null
   * or is not a properly-formatted double
   * @param va string to be converted
   * @return value of the string as a double or 0 if anything goes
   * wrong.
   */
  public static double doubleFromString(String va) {
    try {
      return Double.valueOf(va);
    } catch (Exception e) { /* */ }
    return 0;
  }

  /**
   * Return the value of a parameter as a double.  Returns 0 if the
   * parameter is not found or is not a properly-formatted double
   * @param param name of the desired parameter
   * @return value of the parameter as a double or 0 if it is not a
   * proper double
   */
  public double getDouble(String param) {
    return doubleFromString((String)get(param));
  }

  /**
   * Convert a string to a long, returning 0 if the string is null or
   * is not a properly-formatted long
   * @param va string to be converted
   * @return value of the string as a long or 0 if anything goes wrong.
   */
  public static long longFromString(String va) {
    try {
      return Long.decode(va);
    } catch (Exception e) { /* */ }
    return 0;
  }

  /**
   * Convert a string to a long, returning null if the string is null
   * or if is not a properly-formatted long
   * @param va string to be converted
   * @return value of the string as a long or 0 if anything goes wrong.
   */
  public static Long longNullFromString(String va) {
    if (va == null) { return null; }
    try {
      return Long.decode(va);
    } catch (Exception e) { /* */ }
    return null;
  }
  /**
   * Convert a string which might have a period in it to a long,
   * returning 0 if the string is null or is not a properly-formatted
   * long.  This returns the integer valule of the string.
   * @param va string to be converted
   * @return value of the string as a long or 0 if anything goes wrong.
   */
  public static long longFromPunctString(String va) {
    if ((va == null) || (va.length() <= 0)) { return 0; }
    try {
      int ix;
      if ( (ix = va.indexOf(".")) >= 0) {
	if (ix == 0) {
	  return 0;
	}
	  return Long.valueOf(va.substring(0,ix));
      }
    } catch (Exception e) { /* */ }
    return 0;
  }

  /**
   * Return the value of a parameter as a long, returns 0 if the
   * parameter is not found or it is not a properly-formatted long
   * @param param name of the parameter to be converted to long
   * @return value of the parameter as a long, returns 0 if the
   * parameter is not found
   */
  public long getLong(String param) {
    return longFromString((String)get(param));
  }

  /** See if a Long is nonzero, null is OK, forces false
   *@param l the long to be tested
   *@return true if the long is nonzero*/
  public static boolean longIsNonzero(Long l) {
    return ((l != null) && (l.longValue() != 0));
  }

  /** See if a Long is positive, null is OK, forces false
   *@param l the long to be tested
   *@return true if the long value is > 0*/
  public static boolean longIsPositive(Long l) {
    return ((l != null) && (l.longValue() > 0));
  }

  /**
   * Prepare a string for comparison with a string which is in standard format
   * @param in input string
   * @return trimmed, and with all redundant spaces removed.  Multiple
   * spaces are reduced to one space.
   */
  public static String canonize(String in) {
    in = in.trim().replace("  ", " ");
    while (in.indexOf("  ") >= 0) { in = in.replace("  ", " "); }
    return in;
  }

  /** Convert a string to initial caps, which is sometimes known as
   * camel-case, also collapses successive white space characters to one
   * space.
   * @param input string to be converted
   * @return String with redundant spaces removed and all words given initial
   * capitals*/
  public static String initCaps(String input) {
    if (input == null) { return null; }
    // Note that there is no trailing or leading white space due to the trim
    StringBuilder sb = new StringBuilder(input.toLowerCase().trim());
    boolean hadSpace = true;
    char ch;
    int i;

    for (i=0; i<sb.length(); i++) {
      if (hadSpace) {
        if (Character.isLowerCase(ch = sb.charAt(i))) {
          sb.setCharAt(i, Character.toUpperCase(ch));
          hadSpace = false;
        } else if (Character.isWhitespace(ch)) {
          sb.deleteCharAt(i);
          while (Character.isWhitespace(sb.charAt(i))) {
            sb.deleteCharAt(i);
          }
          if (Character.isLowerCase(ch = sb.charAt(i))) {
            sb.setCharAt(i, Character.toUpperCase(ch));
            hadSpace = false;
          }
        } else {
          hadSpace = false;
        }
      } else {
        if (Character.isWhitespace(ch = sb.charAt(i)) || (ch == '.')) {
          hadSpace = true;
        }
      }
    }
    return sb.toString();
  }

  /** See if a string has at least one character in it.
   *@param s the string to be tested
   *@return true if the string is on-null and has characters*/
  public static boolean stringHasContent(String s) {
    return ((s != null) && (s.length() > 0));
  }

  /** See if a string has at least one non-space character in it.  Length
   * is assessed after trimming the string.
   *@param s the string to be tested
   *@return true if the string is on-null and has characters*/
  public static boolean stringHasNonspace(String s) {
    return ((s != null) && (s.trim().length() > 0));
  }
}
