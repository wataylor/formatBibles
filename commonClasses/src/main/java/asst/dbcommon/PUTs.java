/* @name PUTs.java

   Copyright (c) 2002-2012 Advanced Systems and Software Technologies,
   LLC (All Rights Reserved)

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

package asst.dbcommon;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.List;
import java.util.ArrayList;
import java.util.UUID;

/** Utilities for use with a connection pool.  All methods take a
 * connection as an input.  The connection is used but not closed.
 * Any statements and result sets created are closed.</p>
 * <p> This class was written before JPA and will be replaced.
 *
 * @author Material Gain
 * @since 2005 01
 *
 */

public class PUTs {

  static final long serialVersionUID = 1;

  /** Various databases have different commands for getting the next
   * ID from an autoincrement table.  This string specifies how to get
   * the next ID from the current database.*/
  public static String dbaseGetNextValue = "select LAST_INSERT_ID()";

  /** Obligatory constructor.*/
  public PUTs() { /* */ }

  /** Execute any update or insert statement against the database.
      @param query insert or update query
      @param conn database connection which must be open and valid.
      @exception SQLException on SQL error. */
  public static void anyStatement(String query, Connection conn)
    throws SQLException {
    Statement stmt = null;

    // TODO timeout
    try {
      stmt = conn.createStatement();
      anyStatement(query, stmt);
    } finally {
      if (stmt != null) { stmt.close(); }
    }
  }

  /** Execute any update or insert statement against the database.
      @param query insert or update query
      @param stmt database statement which must be open and valid.
      @exception SQLException on SQL error. */
  public static void anyStatement(String query, Statement stmt)
    throws SQLException {
    try {
      stmt.execute(query);
    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  /** Turn the string or strings in a result set of a query into a list
   which is intended to be used to generate a selection list.
   @param query which has been properly formatted for the current database
   @param conn open connection to the database
   @return list with one element per row of results.  If more than
   one column was selected, the column values from a row are
   concatenated and a | is put between the first column and the rest
   of the columns in each row; the code to generate the selection list
   splits the string so that the string before the | is the selection
   value and the rest of the string is displayed.  The list is empty
   if the search has no results.
  @exception SQLException on SQL error.*/
  public static List<String> listQuery(String query, Connection conn)
    throws SQLException {
    List<String> list;		// Results returned to caller
    Statement stmt = null;	// Retrieve all the rows
    StringBuilder value;		// Result of the query
    ResultSet results = null;	// List of all values in the query
    ResultSetMetaData meta = null;
    int columnCount;
    int i;
    /**/

    list  = new ArrayList<String>();
    value = new StringBuilder();

    try {
      stmt    = conn.createStatement();
      results = stmt.executeQuery(query);
      meta    = results.getMetaData();
      columnCount = meta.getColumnCount();
      while (results.next()) {
	value.setLength(0);
	for (i=1; i<=columnCount; i++) {
	  if (i == 2) { value.append("|"); }
	  if (i > 2)  { value.append(" "); }
	  value.append(results.getString(i));
	}
	if (value.length() > 0) {
	  list.add(value.toString());
	}
      }
    } finally {
      if (results != null) { results.close(); }
      if (stmt    != null) { stmt.close(); }
    }
    return list;
  }

  /** Safely convert a string to an integer, returning 0 in case of
      difficulty.
  @param val input string
  @return integer representation of the string or 0*/
  public static int integerFromString(String val) {
    try {
      return Integer.parseInt(val);
    } catch (Exception e) {
    }
    return 0;
  }

  /** Safely convert a string to a floating point number, returning 0
      in case of difficulty.
  @param val input string
  @return floating point representation of the string or 0*/
  public static float floatFromString(String val) {
    try {
      return Float.parseFloat(val);
    } catch (Exception e) {
    }
    return (float)0.0;
  }

  /** Safely convert a result set string to a non-null string,
      returning a non-null empty string if the result set does not
      have enough strings or if the selected string is null.  The
      effect is to be able to scan a result set without generating a
      SQLException or having to process a null string.
  @param result result set from a prior SQL query.
  @param w selects a string from the result set
  @return the string or the empty string if the column is null or if w
  is out of range for the result set*/
  public static String stringFromResult(ResultSet result, int w) {
    String str = null;
    /**/

    try {
      str = result.getString(w);
    } catch (Exception e) { }
    if (str == null) { str = ""; }
    return str;
  }

  /** Generate the next available value for the ID column of a MySQL
   * table.  This method is synchronized because new values are
   * available only for the same connection and so long as no other
   * request comes via that connection.
   * @param table name of the table whose ID column is to be incremented.
   * @param conn database connection. */
  public static synchronized int nextIDForTable(String table, Connection conn)
    throws SQLException {
    ResultSet results = null;		// Results of the retrieval
    Statement stmt = null;		// Retrieve the sequence number
    String query = null;
    int i;			// The resulting int
    /**/

    i = 0;			// Returned if fails
    try {
      stmt = conn.createStatement();
      // MySQL auto increments ONLY when a 0 or NULL value is put
      // into an autoincrement column.  In effect, MySQL uses the ID
      // column of the table as a sequence.  The insert command
      // inserts default values into all columns other than the ID
      // column
      query = "insert into " + table + " (" + "ID" + ") values (NULL)";
      stmt.execute(query);

      // The LAST_INSERT_ID is good ONLY immediately after an insert
      // into an autoincrement column.  This retrieves the ID which
      // was given to the most recent insert and returns it to the
      // caller so that the caller knows the ID of the most recent
      // addition to the table.
      query   = dbaseGetNextValue;
      results = stmt.executeQuery(query);
      if (results.next()) {	// There were some rows in the table
	i = results.getInt(1);
      } else {		// The table was empty
	i = 1;
      }
      return i;
    } finally {
      if (results != null) { results.close(); }
      if (stmt != null)    { stmt.close(); }
    }
  }

  /** Test a string for content.
   * @param str input string
   * @return true if it is empty, which means either null or zero length
   */
  public static boolean isStringMTP(String str) {
    return ((str == null) || (str.length() <= 0));
  }

  /** Generate a unique random string for use as a database key.
   * These strings are random across all domains so any such string
   * can be guaranteed to be a unique key for use in any database.
   * This method leaves out dashes which the toString() method
   * inserts.  That makes these keys shorter.  The dashes add no
   * information because they are the same in all UUIDs.
   * @return 32 character random string
   */
  public static String anyNewDBKey() {
    UUID uuid = UUID.randomUUID();
    String firstHalf = Long.toHexString(uuid.getMostSignificantBits());
    while (firstHalf.length() < 16)  { firstHalf  = "0" + firstHalf; }
    String secondHalf = Long.toHexString(uuid.getLeastSignificantBits());
    while (secondHalf.length() < 16) { secondHalf = "0" + secondHalf; }
    return firstHalf + secondHalf;
  }

}
