package asst.gentlerKJB;

import java.io.File;
import java.io.FileWriter;
import java.io.PrintWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import asst.common.DescribeArgs;
import asst.common.MainArgs;
import asst.gentlerKJB.utils.PassingItems;
import asst.gentlerKJB.utils.WordUpgradeUtils;
import asst.hssf.SSU;
import asst.hssf.WorkbookManager;

/** Read input files and change words in them as specified by the
 * input dictionary.
 * @author Material Gain
 * @since 2025 12
 */
public class GentlerKBJMain {
  /** Describe the purpose of the command line args for the Help function.  */
  public static Map<String, String> argDescs = new HashMap<String, String>();
  static {
    argDescs.put("help", "If \"+help\" is specified, nothing else is run.");
    argDescs.put("doIt", "\"+doIt\" must be set to take any action."
	+ "  If it is not set, no action is taken and information about"
	+ " actions that would be taken is printed instead.");
    argDescs.put("verbose", "\"+verbose\" increases the printout volume.");
    argDescs.put("inputPath", "File path to a folder where the input files are.");
    argDescs.put("firstFile", "Files in the input path are sorted in alphabetical order."
	+ " Files are skipped until this file is fouknd.");
    argDescs.put("outputPath", "File path to a folder where modified files are written. "
	+ "Must end with a /.");
    argDescs.put("dictionary", "Path to a .xlsx file which tells which words to change.");
    argDescs.put("count", "Tells how many input files to process.");
  }
  /** +help is the default value so that the program explains
   * the parameters if it is called with no arguments. */
  static final String[] DEFAULT_ARGS = {
      "-doIt", "-verbose",
      "inputPath=/Sync/Biblical/asciiBible",
      "outputPath=/temp/KJB/",
      "dictionary=/Sync/Biblical/KJV/Gentle/KJBWordUpdates.xlsx",
      "firstFile=40MAT.TXT",
      "count=40",
      "+help",
  };
 
  /**
   * @param args processed by the MainArgs class. 
   */
  public static void main(String[] args) {
    MainArgs carg = new MainArgs(DEFAULT_ARGS);
    carg.parseArgs(args);

    if (carg.getBoolean("help")) {
      StringBuilder sb =
	  DescribeArgs.describeArgs(carg, DEFAULT_ARGS, argDescs,
	      "There may be NO spaces before or after the = in a string parameter value.\n"
		  + "If a value has spaces, the entire string must be enclosed in quotes or the space must be escaped.\n");
      System.out.print(sb.toString());
      System.exit(0);
    }

    String inputPath = (String)carg.get("inputPath");
    String outputPath = (String)carg.get("outputPath");
    String dictionaryFile = (String)carg.get("dictionary");
    String firstFile = (String)carg.get("firstFile");
    int count = carg.getInt("count");

    WorkbookManager wm = new WorkbookManager();
    wm.fileName = dictionaryFile;
    File file = new File(wm.fileName);

    XSSFWorkbook wb = null;
    int verseCount = 0;
    try {
      if (!file.canRead()) {
	throw new RuntimeException("File " + wm.fileName + " cannot be read.");
      }
      wb = new XSSFWorkbook(file);
      wm.wb = wb;
      if (wm.pickSheet("WordChanges") == null) {
	System.out.println("Spreadsheet has no sheet named WordChanges.");
	System.exit(1);
      }
      wm.makeFormatters();

      Path inputDir = Paths.get(inputPath);
      if (!Files.isDirectory(inputDir)) {
	System.err.println("Input path is not a directory: " + inputDir);
	System.exit(1);
      }

      Path outputDir = Paths.get(outputPath);
      if (!Files.exists(outputDir)) {
	Files.createDirectories(outputDir);
      }

      PrintWriter explanationWriter = new PrintWriter(new FileWriter(new File(outputDir.toFile(), "explanation.txt")));

      int processed = 0;
      List<Path> txtFiles = new ArrayList<>();
      java.util.stream.Stream<Path> stream = Files.list(inputDir);
      try {
	stream.filter(p -> p.toString().toLowerCase().endsWith(".txt"))
	.forEach(txtFiles::add);
      } finally {
	stream.close();
      }

      txtFiles.sort((a, b) -> a.getFileName().toString().compareTo(b.getFileName().toString()));

      boolean foundFirst = firstFile == null || firstFile.isEmpty();

      for (Path inputFile : txtFiles) {
	String fileName = inputFile.getFileName().toString();
	if (!foundFirst) {
	  if (fileName.equals(firstFile)) {
	    foundFirst = true;
	  } else {
	    continue;
	  }
	}

	if (processed >= count) {
	  break;
	}

	try {
	  List<String> lines = Files.readAllLines(inputFile, StandardCharsets.UTF_8);
	  List<String> out = new ArrayList<>(lines.size());
	  for (String line : lines) {
	    out.add(upgradeLine(fileName.substring(0, 2), line, wm));
	    verseCount++;
	  }

	  Path outputFile = outputDir.resolve(inputFile.getFileName());
	  Files.write(outputFile, out, StandardCharsets.UTF_8);
	  System.out.println("Processed: " + inputFile + " -> " + outputFile);
	  processed++;
	} catch (Exception e) {
	  System.out.println("ERR processing " + inputFile + ": " + e.getMessage());
	  e.printStackTrace();
	  System.exit(1);
	}
      }
      explanationWriter.println("#" + processed + " files processed,"
	  + " " + WordUpgradeUtils.verseChanges.size()
	  + " verses changed out of " + verseCount + ".");
      for (String v : WordUpgradeUtils.verseChanges) {
	explanationWriter.print(v + "_");
      }
      explanationWriter.println("\n");

      // Get keys in order
      Set<String> keys = WordUpgradeUtils.wordChanges.keySet();
      for (String key : keys) {
	explanationWriter.print(key + ": ");
	StringBuilder sb = WordUpgradeUtils.wordChanges.get(key);
	if (sb != null) {
	  explanationWriter.println(sb.toString());
	  if (sb.length() > 5) {
	    explanationWriter.println();
	  }
	}
      }
      explanationWriter.close();
      System.out.println("Finished processing.");
    } catch (Exception e) {
      System.out.println("ERR accessing directories: " + e.getMessage());
      System.exit(1);
    } finally {
      try {
	if (wb != null) { wb.close(); }
      } catch (Exception e) {
	System.out.println("ERR closing work book " + e.getMessage());
      }
    }
  }

  /** Given one verse, upgrade it by reading through the spreadsheet
   * and changing words as called for by the spreadsheet
   * @param bkno 2-digit book number
   * @param line original verse
   * @param wm workbook manager
   * @return modified line of text
   */
  public static String upgradeLine(String bkno, String line, WorkbookManager wm) {
    PassingItems pi = new PassingItems(line, "w", "w");
    pi.bkno = bkno;
    for (int i = wm.sheet.getFirstRowNum(); i <= wm.sheet.getLastRowNum(); i++) {
      Row row = wm.sheet.getRow(i);
      String oldWord = SSU.getFormattedCell(0, row);
      if ((oldWord == null) || oldWord.startsWith("#")) { continue; }
      String newWord = SSU.getFormattedCell(1, row);
      String verb = SSU.getFormattedCell(2, row);
      if ((newWord == null) || (newWord.length() <= 0)) { break; }
      pi.setWords(oldWord, newWord);
      if ((verb != null) && ("Not mark".equals(verb))) {
	WordUpgradeUtils.replaceWord(pi);
      } else if ((verb != null) && verb.startsWith("Only in")) {
	/* Modernize the word only in specified verses */
	if (verb.indexOf(pi.bookChapVerse) > 0) {
	  WordUpgradeUtils.modernizeWord(pi);
	}
      } else {
	WordUpgradeUtils.modernizeWord(pi);
      }
    }
    return pi.getEditedLine();
  }
}
