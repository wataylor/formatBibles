/* @name TwixtPoiPojo.java

   Copyright (c) 2002-2013 Advanced Systems and Software Technologies, LLC
   (All Rights Reserved)

-------- Licensed Software Proprietary Information Notice -------------

This software is a working embodiment of certain trade secrets of
AS&ST.  The software is licensed only for the day-to-day business use
of the licensee.  Use of this software for reverse engineering,
decompilation, use as a guide for the design of a competitive product,
or any other use not for day-to-day business use is strictly
prohibited.

All screens and their formats, color combinations, layouts, and
organization are proprietary to and copyrighted by AS&ST.

All rights are reserved.

Authorized AS&ST customer use of this software is subject to the terms
and conditions of the software license executed between Customer and
AS&ST.

------------------------------------------------------------------------

*/

package asst.hssf;

import java.lang.reflect.Field;
import java.lang.reflect.Type;
import java.util.Date;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import asst.dbcommon.AColumn;
import asst.dbcommon.ASSColumn;
import asst.dbcommon.PUTs;

/** Utilities to convert between Spread Sheet rows and annotated POJOs
 *
 * @author Material Gain
 * @since 2014 06
 *
 */

public class TwixtPoiPojo {

  static final long serialVersionUID = 1;

  /** Obligatory constructor.*/
  public TwixtPoiPojo() { /* */ }

  /**
   * Accumulate a complaint into a text string which is suitable for an
   * alert.
   * @param sb accumulator
   * @param whinge text of the current complaint
   */
  public static void whingeSB(StringBuilder sb, String whinge) {
    if (sb.length() > 0) { sb.append("<br>\n"); }
    sb.append(whinge);
  }

  /**
   * Get information about the spread sheet column which may be associated
   * with a POJO field.  If there is an AColumn annotation, that is used.
   * That annotation may have an sSColName but it always has a columnName.
   * If that annotation does not exist, check for ASSColumn.  This notation
   * must have an sSColName to support spread sheet columns which do not
   * match database columns.
   * @param fld the field which might be annotated
   * @param sb accumulator for error messages
   * @param clazz the class definition for error messages
   * @return column information object or null
   */
  public static TwixtColProps getSpreadSheetColumnName(Field fld,
      StringBuilder sb, Class<?> clazz) {
    String sSColName;
    AColumn aCol;
    ASSColumn aSSCol;
    boolean writeZAN;

    if ( (aCol = fld.getAnnotation(AColumn.class)) != null) {
      sSColName = aCol.sSColName();
      writeZAN = aCol.writeZeroAsNull();
      if (PUTs.isStringMTP(sSColName)) { sSColName = aCol.columnName(); }
      if (PUTs.isStringMTP(sSColName)) {
	whingeSB(sb, "Class " + clazz.getName() + " field " +
	    fld.getName() + " has no database column name.");
	return null;
      }
    } else if ( (aSSCol = fld.getAnnotation(ASSColumn.class)) != null) {
      sSColName = aSSCol.sSColName();
      writeZAN = aSSCol.writeZeroAsNull();
      if (PUTs.isStringMTP(sSColName)) {
	whingeSB(sb, "Class " + clazz.getName() + " field " +
	    fld.getName() + " has no spread sheet column name.");
	return null;
      }
    } else {
      return null;
    }
    TwixtColProps tcp = new TwixtColProps();
    tcp.columnName = sSColName;
    tcp.writeZeroAsNull = writeZAN;
    return tcp;
  }

  /**
   * Fill in an annotated object field values from a spread sheet row.
   * The spread sheet column names must match the database column names
   * in the annotated object
   * @param o the annotated object
   * @param row one spread sheet row
   * @param columnMap map column names to column numbers
   * @return non-empty string if there are errors, otherwise the empty
   * string
   */
  public static String pojoValuesFromSheetRow(Object o, Row row,
					      Map<String, Integer> columnMap) {
    StringBuilder sb = new StringBuilder();
    Class<?> clazz = o.getClass();
    Cell cell;
    Integer sSColNumber;
    String cellStringValue;
    TwixtColProps tcp;
    boolean isNumeric; // What Excel thinks about the cell
    double cellNumberValue;

    for (Field fld : clazz.getFields()) {
      if ( (tcp = getSpreadSheetColumnName(fld, sb, clazz)) == null) {
	continue;
      }
      if ( (sSColNumber = columnMap.get(tcp.columnName)) == null) {
/*	whingeSB(sb, "Class " + clazz.getName() + " field " +
	fld.getName() + " has no column in the spread sheet.");
*/	continue;
      }
      cellNumberValue = 0;
      cellStringValue = "";
      isNumeric = false;
      cell = null;
      try {
	cell = row.getCell(sSColNumber.intValue());
      } catch (Exception e) {
	whingeSB(sb, "Class " + clazz.getName() + " field " +
		 fld.getName() + " has no column " + sSColNumber +
		 " in the spread sheet row.");
	continue;
      }
      if (cell != null) {
	CellType cellType = cell.getCellType();
	if (cellType == CellType.NUMERIC) {
	  cellNumberValue = cell.getNumericCellValue();
	  isNumeric = true;
	} else {
	  isNumeric = false;
	  /* Unfortunately, the POI utilities do not have a type value to
	   * indicate a date type column.  Empty cells return null when
	   * attempts are made to convert them to dates; non-date cells
	   * cause exceptions.*/
	  Date date = null;
	  try {
	    if ( (date = cell.getDateCellValue()) != null) {
	      cellStringValue = SSU.EXCEL_DATE.format(date);
	    }
	  } catch (Exception e) {
	    /* The cell is not a date cell, treat it as a string.*/
	    cellStringValue = cell.getRichStringCellValue().getString().trim();
	  }
	}
      }
      /* One way or another, we have a numeric cell value which defaults to 0
       * and a string cell value which defaults to the empty string.  Set
       * the field: */
      Type type = fld.getGenericType();
      try {
	if (type == String.class) {
	  /* Excel may have declared the cell to be numeric but we want
	   * a string instead.  Format it as a string. */
	  if (isNumeric) { cellStringValue = SSU.df.formatCellValue(cell); }
	  if (tcp.writeZeroAsNull && PUTs.isStringMTP(cellStringValue)) {
	    cellStringValue = "";
	  }
	  fld.set(o, cellStringValue);
	} else {
	  if (!PUTs.isStringMTP(cellStringValue)) { // excel said it was string
	    try { // try to convert the string to a double we can use
	      cellNumberValue = Double.valueOf(cellStringValue);
	    } catch (Exception e) { cellNumberValue = 0; }
	  }
	  if (type == Integer.TYPE || type == Integer.class) {
	    fld.setInt(o, (int)cellNumberValue);
	  } else if (type == Long.TYPE | type == Long.class) {
	    fld.setLong(o, (long)cellNumberValue);
	  } else if (type == Float.TYPE || type == Float.class) {
	    fld.setFloat(o, (float)cellNumberValue);
	  } else if (type == Double.TYPE || type == Double.class) {
	    fld.setDouble(o, cellNumberValue);
	  } else {
	    whingeSB(sb, "Class " + clazz.getName() + " field " +
		     fld.getName() + " has an unexpected type " + type);
	  }
	}
      } catch (Exception e) {
	whingeSB(sb,  "Class " + clazz.getName() + " field " +
		 fld.getName() + " error setting value " + e.toString());
      }
    }
    return sb.toString();
  }

  /**
   * Fill in a spread sheet row from annotated object field values.
   * The spread sheet column names must match the database column names
   * in the annotated object
   * @param o the annotated object
   * @param row one spread sheet row
   * @param columnMap map column names to column numbers
   * @return non-empty string if there are errors, otherwise the empty
   * string
   */
  public static String sheetRowFromPOJOValues(Object o, Row row,
					      Map<String, Integer> columnMap) {
    StringBuilder sb = new StringBuilder();
    Class<?> clazz = o.getClass();
    Cell cell;
    Integer sSColNumber;
    TwixtColProps tcp;
    Object objectFieldValue;

    for (Field fld : clazz.getFields()) {
      if ( (tcp = getSpreadSheetColumnName(fld, sb, clazz)) == null) {
	continue;
      }
      if ( (sSColNumber = columnMap.get(tcp.columnName)) == null) {
/*	whingeSB(sb, "Class " + clazz.getName() + " field " +
	fld.getName() + " has no column in the spread sheet.");
*/	continue;
      }
      /* Will put the field value into a spread sheet cell in the row. */
      Type type = fld.getGenericType();
      try {
	cell = row.getCell(sSColNumber, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
	if (type == String.class) {
	  objectFieldValue = fld.get(o);
	  cell.setCellType(CellType.STRING);
	  cell.setCellValue((objectFieldValue == null) ? "" : objectFieldValue.toString());
	} else if (type == Integer.TYPE || type == Integer.class) {
	  cell.setCellType(CellType.NUMERIC);
	  cell.setCellValue(fld.getInt(o));
	} else if (type == Long.TYPE || type == Long.class) {
	  cell.setCellType(CellType.NUMERIC);
	  cell.setCellValue(fld.getLong(o));
	} else if (type == Float.TYPE || type == Float.class) {
	  cell.setCellType(CellType.NUMERIC);
	  cell.setCellValue(fld.getFloat(o));
	} else if (type == Double.TYPE || type == Double.class) {
	  cell.setCellType(CellType.NUMERIC);
	  cell.setCellValue(fld.getDouble(o));
	} else {
	  whingeSB(sb, "Class " + clazz.getName() + " field " +
		   fld.getName() + " has an unexpected type " + type);
	}
      } catch (Exception e) {
	whingeSB(sb,  "Class " + clazz.getName() + " field " +
		 fld.getName() + " error setting value " + e.toString());
      }
    }
    return sb.toString();
  }
}
