package asst.formatWord.snippets;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.Map;

import asst.common.DescribeArgs;
import asst.common.MainArgs;

/** Read a Bible expressed in a text file and convert it to
 * 66 chapter files in the format expected by the gentlerKJB and
 * formatWord programs.
 * @author Material gain
 * @since 2026 02
 * 
 */
public class TxtKJBTo66ChapterFilesMain {

  /** Describe the purpose of the command line args for the Help function.  */
  public static Map<String, String> argDescs = new HashMap<String, String>();

  static {
    argDescs.put("help", "If \"+help\" is specified, nothing else is run."
        + " Enter -help to turn help off to run the program.");
    argDescs.put("doIt", "\"+doIt\" must be set to take any action."
	+ "  If it is not set, no action is taken and information about"
	+ " actions that would be taken is printed instead.");
    argDescs.put("verbose", "\"+verbose\" increases the printout volume.");
    argDescs.put("inputFile", "Text file to be split ito 66 chapter files.");
    argDescs.put("outputPath", "Path where the converted files are written. ");
    argDescs.put("count", "Tells how many books to process.");
  }

  /** +help is the default value so that the program explains
   * the parameters if it is called with no arguments. */
  static final String[] DEFAULT_ARGS = {
      "-doIt", "-verbose",
      "inputFile=/sync/Biblical/KJV/KJB-PCE-127-With.txt",
      "outputPath=/temp/KJB/kjv-bibleprotector-com",
      "count=67", // Stop after 67 chapters
      "+help",
  };

  /**
   * @param args as described in the argDescs map
   */
  public static void main(String[] args) {
    MainArgs carg = new MainArgs(DEFAULT_ARGS);
    carg.parseArgs(args); // updates or overrides the default values
    int bookCount = 0; /* number of books to be processed */;
    int bookNum = 0; /* On this book, 0 is before any book. */
    
    String inputFile = null;
    String outputPath = null;
    String line = null;
    String outputFile = null;
    String inputLineStart = "  "; /* Characters at beginning of an input line. */
    String lineHead = "     "; /* Characters written at the beginning of an output line. */
    BufferedReader reader = null;
    BufferedWriter writer = null;

    if (carg.getBoolean("help")) {
      StringBuilder sb =
	  DescribeArgs.describeArgs(carg, DEFAULT_ARGS, argDescs,
	      "There may be NO spaces before or after the = in a string parameter value.\n"
		  + "If a value has spaces, the entire string must be enclosed in quotes or the space must be escaped.\n");
      System.out.print(sb.toString());
      System.exit(0);
    }

    inputFile = (String)carg.get("inputFile");
    outputPath = (String)carg.get("outputPath");
    bookCount = carg.getInt("count");
    int isp = 0; /* Index of ht first space in the line */

    try {
      reader = new BufferedReader(
	  new InputStreamReader(
	      new FileInputStream(inputFile),
	      StandardCharsets.UTF_8));
      while ((line = reader.readLine()) != null) {
	if (!line.startsWith(inputLineStart)) {
	  isp = line.indexOf(" ");
	  inputLineStart = line.substring(0, isp);
	  // Close the current output file if there is one
	  if (writer != null) {writer.close(); }
	  // Determine the next output filename
	  outputFile = BOOK_FILE_NAMES[++bookNum];
	  if (bookNum > bookCount) { break; }
	  // Open the new output file in the same folder
	  writer = new BufferedWriter(
	      new OutputStreamWriter(
	          new FileOutputStream(outputPath + "/" + outputFile),
	          StandardCharsets.UTF_8));
	  lineHead = outputFile.substring(2, 5);
	}
	int ix = line.indexOf(" ", isp + 1); /* Need 2 spaces after the verse number */
	String remainder = line.substring(ix + 1);
	remainder = remainder.replace("[", "<i>");
	remainder = remainder.replace("]", "</i>");
	writer.write(lineHead + line.substring(isp, ix) + "  " + remainder);
	writer.newLine();
      }
    } catch (Exception e) {
      System.err.println("ERR input file: " + inputFile + " " + outputFile);
      e.printStackTrace();
    }

    try {
      // After the loop, close the last output file
      if (writer != null) {
	writer.close();
      }
      reader.close();

    } catch (Exception e) {
      System.err.println("ERR output file " + outputPath + " " + outputFile);
      e.printStackTrace();
    }
    System.out.println("Wrote " + outputFile);
  }

  /** Names of the files to be written for each book of the Bible */
  public static final String[] BOOK_FILE_NAMES = {
    "",
    "01GEN.TXT",
    "02EXO.TXT",
    "03LEV.TXT",
    "04NUM.TXT",
    "05DEU.TXT",
    "06JOS.TXT",
    "07JDG.TXT",
    "08RTH.TXT",
    "09SA1.TXT",
    "10SA2.TXT",
    "11KI1.TXT",
    "12KI2.TXT",
    "13CH1.TXT",
    "14CH2.TXT",
    "15EZR.TXT",
    "16NEH.TXT",
    "17EST.TXT",
    "18JOB.TXT",
    "19PSA.TXT",
    "20PRO.TXT",
    "21ECC.TXT",
    "22SON.TXT",
    "23ISA.TXT",
    "24JER.TXT",
    "25LAM.TXT",
    "26EZE.TXT",
    "27DAN.TXT",
    "28HOS.TXT",
    "29JOE.TXT",
    "30AMO.TXT",
    "31OBA.TXT",
    "32JON.TXT",
    "33MIC.TXT",
    "34NAH.TXT",
    "35HAB.TXT",
    "36ZEP.TXT",
    "37HAG.TXT",
    "38ZEC.TXT",
    "39MAL.TXT",
    "40MAT.TXT",
    "41MAR.TXT",
    "42LUK.TXT",
    "43JOH.TXT",
    "44ACT.TXT",
    "45ROM.TXT",
    "46CO1.TXT",
    "47CO2.TXT",
    "48GAL.TXT",
    "49EPH.TXT",
    "50PHI.TXT",
    "51COL.TXT",
    "52TH1.TXT",
    "53TH2.TXT",
    "54TI1.TXT",
    "55TI2.TXT",
    "56TIT.TXT",
    "57PHM.TXT",
    "58HEB.TXT",
    "59JAM.TXT",
    "60PE1.TXT",
    "61PE2.TXT",
    "62JO1.TXT",
    "63JO2.TXT",
    "64JO3.TXT",
    "65JUD.TXT",
    "66REV.TXT",
  };
}
