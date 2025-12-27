package asst.hssf;

import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Convert result sets to work sheets
 * @author Material Gain
 * @since 2015 03
 */
public class ResultSetToSheet {

  /**
   * Create or overwrite a work sheet in the current spread sheet
   * @param sheetName identifies the sheet to create
   * @param rs result set
   * @param wb work book
   * @return the created sheet
   * @throws Exception when things go wrong
   */
  public static Sheet resultSetToWBSheet(String sheetName, ResultSet rs, Workbook wb)
      throws Exception {
    ResultSetMetaData meta = rs.getMetaData();
    String columns[] = new String[meta.getColumnCount()];
    for (int i = 1; i<=meta.getColumnCount(); i++) {
      columns[i-1] = meta.getColumnLabel(i);
    }
    Sheet sheet = SSU.getOrMakeSheet(sheetName, false, columns, wb);
    while (rs.next()) {
      for (int i = 1; i<=meta.getColumnCount(); i++) {
	columns[i-1] = rs.getString(i);
      }
      SSU.printlnSheet(sheet, columns);
    }
    return sheet;
  }
}
