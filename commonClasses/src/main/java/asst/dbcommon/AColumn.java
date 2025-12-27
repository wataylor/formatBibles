package asst.dbcommon;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Specify information about a column name and field for a table row.
 * The annotated class must have fields whose data types match the column
 * names which are returned by the select.</p>
 * <p> This annotation was written before JPA appeared and is obsolete.  It is
 * used to map spread sheet column names to Java fields.
 * @author Material Gain
 * @since 2014 02
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface AColumn {
  /** Names the table column */
  public String columnName() default "";
  /** Names a spread sheet column for reading from excel if it is different
   * from the database column name*/
  public String sSColName() default "";
  /** The primary key column name must agree with the column name specified
   * in the ATable annotation for the class. */
  public boolean primaryKey() default false;
  /** If > 0, sets the maximum length of a string to be written.  It does
   * not affect any other field types.*/
  public int maxColWidth() default 0;
  /** If this is true, an empty string and a zero numerical value are not
   * written to the database in an update statement.*/
  public boolean notWriteEmpty() default false;
  /** Numerical zero values and empty strings are written as SQL nulls.
   * SQL nulls are always read as zero numerical values; this flag means
   * that null string values are read as the empty string. */
  public boolean writeZeroAsNull() default false;
  /** Ignore comparisons of this field when generating update statements.
   * This is used for parameters such as a created time which should
   * never change.*/
  public boolean ignoreUpdate() default false;
}
