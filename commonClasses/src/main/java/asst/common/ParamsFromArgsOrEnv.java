package asst.common;

import java.util.HashMap;
import java.util.Map;

/** read data from the command line arguments or from the environment and make sure that
 * all required parameters are supplied.  Command line values override values in the
 * environment or supply values that are missing from the environment.
 * @author Material Gain
 * @since 2024 09
 */
public class ParamsFromArgsOrEnv {

  /**  This holds parameter values keyed by the command line arg name.  This is a global 
   * variable so that any part of the application can get whatever parameters it needs.*/
  public static Map<String, String> PARAM_MAP = new HashMap<String,String>();

  // This will hold a list of any missing parameters
  StringBuilder missingArgs = new StringBuilder();

  /** Obligatory constructor.*/
  public ParamsFromArgsOrEnv () { /* */ }

  /** Analyze the command line parameters and the environment variables to make sure that
   * all required information is present.  Add values from the environment to the list of
   * command-line arguments.
   * @param carg Command-line parameters
   * @param requirements list of data needed for the program to run at all.  Data from
   * a command line parameter is used before 
   */
  public ParamsFromArgsOrEnv(MainArgs carg, String[] requirements) {
    String value;
    for (String arg : requirements ) {
      value = (String)carg.get(arg);
      if ((value != null) && (value.length() > 0)) {
	// System.out.print(arg + " " + value + "  ");
	PARAM_MAP.put(arg, value);
      } else if ((((value = System.getenv(arg))) != null) && (value.length() > 0)) {
	PARAM_MAP.put(arg, value);
	carg.put(arg, value);
      } else {
	if (missingArgs.length() > 0) {
	  missingArgs.append(", ");
	}
	missingArgs.append(arg);
      }
    }
    // System.out.println();
  }

  /**
   * @return true if an argument in the required list was missing.
   */
  public boolean getIsError() {
    return missingArgs.length() > 0;
  }

  /**
   * @return comma-separated human-friendly list of missing arguments
   */
  public String getMissingArgs() {
    return missingArgs.toString();
  }

  /** Empty the map for another use */
  public void emptyMap() {
    PARAM_MAP.clear();
  }

  /**  Return the value of a parameter.  This is made a separate static method so that
   * it can be changed.  Future implementations may store parameters in other ways, perhaps
   * in a cloud secret store
   * @param key Desired parameter
   * @return parameter value
   */
  public static String getFromMap(String key) {
    return PARAM_MAP.get(key);
  }
}
