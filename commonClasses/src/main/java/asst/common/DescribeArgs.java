/* @name DescribeArgs.java

   Copyright (c) 2002-2025 Advanced Systems and Software Technologies, LLC
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

import java.util.Map;

/** Describe command line arguments
 * @author Material Gain
 * @since 2025 04
 */

public class DescribeArgs {

  /**
   * @param carg Hash table of command line or environment variables
   * @param defaultArgs All default argument values for the service
   * @param argDescs Map of argument names to descriptions
   * @param why String to go before the argument descriptions
   * @return Printable explanations
   */
  public static StringBuilder describeArgs(MainArgs carg, String[] defaultArgs,
      Map<String, String> argDescs, String why) {
    StringBuilder sb = new StringBuilder("\n");
    if ((why != null) && (why.length() > 0)) {
      sb.append(why);
    }
    String key;
    int ix;
    for (String arg : defaultArgs) {
      if (arg.startsWith("DBPASS")) {
	sb.append("DBPASS=<redacted>");
      } else {
	sb.append(arg);
      }
      key = ((arg.startsWith("+") || arg.startsWith("-")) ? arg.substring(1) : arg);
      ix = (key.indexOf("="));
      if (ix > 0) {
	key = key.substring(0, ix);
      }
      if ((key = argDescs.get(key)) == null) {
	sb.append("\n");
      } else {
	sb.append("\t" + key + (((key.length() > 70) ? "\n" : "")) + "\n");
      }
    }
    return sb;
  }
}
