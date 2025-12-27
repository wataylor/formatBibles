package asst.dbcommon;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Specify information about a spread sheet column name for a sheet row
 * The annotated class must have fields whose data types match the column
 * values in the row.  This annotation is used only if a field does not
 * have an AColumn annotation.  AColumn is for fields which appear both
 * in the spread sheet and are written to columns in the database.  This
 * is for data which appear only in the spread sheet and not in the
 * database or where the spread sheet column has a different name.</p>
 * <p>It is examined only for fields which do not have AColumn.  If
 * a field has AColumn, the values in that annotation are used and this one
 * is ignored.
 * @author Material Gain
 * @since 2014 12
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ASSColumn {
  /** Names a spread sheet column for reading from excel.  This is used
   * if the excel column does not match a database column.*/
  public String sSColName() default "";
  /** Ignore comparisons of this field when generating row updates.
   * This is used for parameters such as a created time which should
   * never change.*/
  public boolean ignoreUpdate() default false;
  /** Numerical zero values and empty strings are written as SQL nulls.
   * SQL nulls are always read as zero numerical values; this flag means
   * that null string values are read as the empty string. */
  public boolean writeZeroAsNull() default false;
}
