package asst.formatWord;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.VerticalAlign;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColumns;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STSectionMark;

import asst.common.DescribeArgs;
import asst.common.MainArgs;
import asst.hssf.WorkbookManager;

/** Read input files and change words in them as specified by the
 * input dictionary.
 * @author Material Gain
 * @since 2025 12
 */
public class FormatWordMain {
  /** Describe the purpose of the command line args for the Help function.  */
  public static Map<String, String> argDescs = new HashMap<String, String>();

  // Page size and margin constants (in twentieths of a point)
  // 6 x 9 inches
  private static final BigInteger PAGE_WIDTH = BigInteger.valueOf(8640);   // 6 inches (8.5" = 12240)
  private static final BigInteger PAGE_HEIGHT = BigInteger.valueOf(12960); // 9 inches (11" = 15840)
  private static final BigInteger MARGIN_TOP = BigInteger.valueOf(720);    // 0.5 inch
  private static final BigInteger MARGIN_BOTTOM = BigInteger.valueOf(720); // 0.5 inch
  private static final BigInteger MARGIN_LEFT = BigInteger.valueOf(720);   // 0.5 inch
  private static final BigInteger MARGIN_RIGHT = BigInteger.valueOf(720);  // 0.5 inch
  private static final BigInteger MARGIN_HEADER = BigInteger.valueOf(720);  // 0.5 inch
  private static final BigInteger MARGIN_FOOTER = BigInteger.valueOf(720);  // 0.5 inch

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
    argDescs.put("templateFile", "Path to a .docx template file with predefined styles."
	+ " The generated paragraphs are put at the end of this file.");
  }
  /** +help is the default value so that the program explains
   * the parameters if it is called with no arguments. */
  static final String[] DEFAULT_ARGS = {
      "-doIt", "-verbose",
      "inputPath=/temp/KJB/",
      "outputPath=/temp/KJB/",
      "dictionary=/Sync/Biblical/KJV/Gentle/KJBWordUpdates.xlsx",
      "templateFile=/Sync/Biblical/KJV/Gentle/GentleKJBNT.docx",
      "firstFile=40MAT.TXT",
      "count=50",
      "+help",
  };

  /**  List of file names which are to be ignored.
   */
  public static Set<String> skip_files = new HashSet<String>();
  static {
    skip_files.add("explanation.txt");
  }

  public static final String [] needed_sheets = {
      "WordChanges", "BookNames", "Footnotes",
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
    String templateFile = (String)carg.get("templateFile");
    String firstFile = (String)carg.get("firstFile");
    int count = carg.getInt("count");


    Path outputPlace = Paths.get(outputPath);
    if (!Files.exists(outputPlace)) {
	try {
	  Files.createDirectories(outputPlace);
	} catch (IOException e) {
	  e.printStackTrace();
	  System.exit(-1);
	}
    }

    // Check if output directory is writable
    File outputDir = new File(outputPath);
    if (!outputDir.exists()) {
      System.err.println("Output directory does not exist: " + outputPath);
      System.exit(1);
    }
    if (!outputDir.isDirectory()) {
      System.err.println("Output path is not a directory: " + outputPath);
      System.exit(1);
    }
    if (!outputDir.canWrite()) {
      System.err.println("Output directory is not writable: " + outputPath);
      System.exit(1);
    }

    // Test that we can write the output file
    File outputFile = new File(outputDir, "Injected.docx");
    if (outputFile.exists() && !outputFile.canWrite()) {
      System.err.println("Output file exists but cannot be written: " + outputFile.getAbsolutePath());
      System.exit(1);
    }

    WorkbookManager wm = new WorkbookManager();
    wm.fileName = dictionaryFile;
    File file = new File(wm.fileName);

    XSSFWorkbook wb = null;
    int verseCount = 0;
    try {
      if (!file.canRead()) {
	throw new RuntimeException("File " + wm.fileName + " cannot be read.");
      }
      /* This reads the entire sheet into memory.  It becomes effectively
       * a RAM cache. */
      wb = new XSSFWorkbook(file);
      wm.wb = wb;
      for (String sname : needed_sheets) {
	if (wm.pickSheet(sname) == null) {
	  System.out.println("Spreadsheet has no sheet named " + sname + ".");
	  System.exit(1);
	}
      }
      wm.makeFormatters();

      Path inputDir = Paths.get(inputPath);
      if (!Files.isDirectory(inputDir)) {
	System.err.println("Input path is not a directory: " + inputDir);
	System.exit(1);
      }

      // PrintWriter explanationWriter = new PrintWriter(new FileWriter(new File(outputDir.toFile(), "explanation.txt")));

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

      // Open template document to preserve styles
      File templateFileObj = new File(templateFile);
      if (!templateFileObj.canRead()) {
	throw new RuntimeException("Template file " + templateFile + " cannot be read.");
      }
      XWPFDocument doc = new XWPFDocument(new FileInputStream(templateFileObj));

      // Add section break to end the template's last section
      XWPFParagraph templateEndPara = doc.createParagraph();
      CTP templateCtp = templateEndPara.getCTP();
      CTSectPr templateSectPr = templateCtp.addNewPPr().addNewSectPr();
      templateSectPr.addNewType().setVal(STSectionMark.CONTINUOUS);

      // Set page numbering format to lowercase Roman numerals
      CTPageNumber pgNum = templateSectPr.addNewPgNumType();
      pgNum.setFmt(org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat.LOWER_ROMAN);

      setPageSizeAndMargins(templateSectPr);

      boolean foundFirst = firstFile == null || firstFile.isEmpty();

      for (Path inputFile : txtFiles) {
	String fileName = inputFile.getFileName().toString();
	if (skip_files.contains(fileName)) { continue; }

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
	  String chapNumSt = fileName.substring(0, 2);
	  int chapNum = Integer.valueOf(chapNumSt);

	  startNextChapter(chapNum, wm, doc);
	  for (String line : lines) {
	    publishVerse(chapNumSt, line, wm, doc);
	    verseCount++;
	  }
	  endTheChapter(chapNum, wm, doc);

	  //Path outputFile = outputDir.resolve(inputFile.getFileName());
	  //Files.write(outputFile, out, StandardCharsets.UTF_8);
	  System.out.println("Processed: " + processed + " verses " + verseCount);
	  processed++;
	} catch (Exception e) {
	  System.out.println("ERR processing " + inputFile + ": " + e.getMessage());
	  e.printStackTrace();
	  System.exit(1);
	}
      }

      // Write and close the document
      try (FileOutputStream out = new FileOutputStream(new File(outputPlace.toFile(), "Injected.docx"))) {
	doc.write(out);
      }
      doc.close();

      // explanationWriter.close();
      System.out.println("Finished processing.");
    } catch (Exception e) {
      System.out.println("ERROR " + e.getMessage());
      System.exit(1);
    } finally {
      try {
	if (wb != null) { wb.close(); }
      } catch (Exception e) {
	System.out.println("ERR closing work book " + e.getMessage());
      }
    }
  }

  /** Apply standard page size and margins to a section.  This has to
   * be done for all sections created. */
  private static void setPageSizeAndMargins(CTSectPr sectPr) {
    // Set page size
    CTPageSz pageSz = sectPr.addNewPgSz();
    pageSz.setW(PAGE_WIDTH);
    pageSz.setH(PAGE_HEIGHT);

    // Set margins
    CTPageMar pageMar = sectPr.addNewPgMar();
    pageMar.setTop(MARGIN_TOP);
    pageMar.setBottom(MARGIN_BOTTOM);
    pageMar.setLeft(MARGIN_LEFT);
    pageMar.setRight(MARGIN_RIGHT);
    pageMar.setHeader(MARGIN_HEADER);
    pageMar.setFooter(MARGIN_FOOTER);
  }

  private static void endTheChapter(int chapNum, WorkbookManager wm, XWPFDocument doc) {
    // End the 2-column section by creating a paragraph with section properties
    XWPFParagraph endSectionPara = doc.createParagraph();
    CTP ctp = endSectionPara.getCTP();
    CTSectPr sectPr = ctp.addNewPPr().addNewSectPr();

    // Set to continuous section break
    sectPr.addNewType().setVal(STSectionMark.CONTINUOUS);

    // Define the 2-column layout for the section that just ended
    CTColumns columns = sectPr.addNewCols();
    columns.setNum(BigInteger.valueOf(2));
    columns.setSpace(BigInteger.valueOf(360)); // 0.25 inch gutter

    // Set page size and margins
    setPageSizeAndMargins(sectPr);
  }

  public static void startNextChapter(int chapNum, WorkbookManager wm, XWPFDocument doc) {
    /* Start the new chapter in a new one-column section.  */
    /* The chapter title is used for the page headers */
    String chapTitle = getChapterTitle(wm, chapNum);
    String chapComment = getChapterIntro(wm, chapNum);

    // Create headers for this section
    XWPFHeaderFooterPolicy policy = doc.createHeaderFooterPolicy();

    // Even page header: tab + chapter title
    XWPFHeader evenHeader = policy.createHeader(XWPFHeaderFooterPolicy.EVEN);
    XWPFParagraph evenPara = evenHeader.createParagraph();
    XWPFRun evenRun = evenPara.createRun();
    evenRun.addTab();
    evenRun.setText(chapTitle);

    // Odd page header: chapter title
    XWPFHeader oddHeader = policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);
    XWPFParagraph oddPara = oddHeader.createParagraph();
    oddPara.createRun().setText(chapTitle);

    // Add chapter title as centered Heading1 paragraph
    XWPFParagraph titlePara = doc.createParagraph();
    titlePara.setStyle("Heading1");

    // Set alignment at the XML level to work with the style
    CTPPr pPr = titlePara.getCTP().isSetPPr() ? titlePara.getCTP().getPPr() : titlePara.getCTP().addNewPPr();
    CTJc jc = pPr.isSetJc() ? pPr.getJc() : pPr.addNewJc();
    jc.setVal(STJc.CENTER);

    titlePara.createRun().setText(chapTitle);

    // Add chapter comment if present
    if (chapComment != null && !chapComment.isEmpty()) {
      XWPFParagraph commentPara = doc.createParagraph();
      commentPara.setStyle("FAH");
      commentPara.createRun().setText(chapComment);
    }

    // Create empty paragraph that will end the 1-column section and start 2-column section
    XWPFParagraph columnBreakPara = doc.createParagraph();
    CTP columnCtp = columnBreakPara.getCTP();
    CTSectPr columnSectPr = columnCtp.addNewPPr().addNewSectPr();

    // This sectPr defines the PREVIOUS section (1-column with headers)
    columnSectPr.addNewType().setVal(STSectionMark.CONTINUOUS);

    // Set to 1 column for the previous section
    CTColumns columns1 = columnSectPr.addNewCols();
    columns1.setNum(BigInteger.valueOf(1));

    // Set page size and margins
    setPageSizeAndMargins(columnSectPr);
  }

  private static String getChapterIntro(WorkbookManager wm, int chapNum) {
    Sheet sheet = wm.wb.getSheet("BookNames");
    Row row = sheet.getRow(chapNum);
    Cell cell = row.getCell(3);
    if (cell == null) { return ""; }
    return cell.getStringCellValue();
  }

  private static String getChapterTitle(WorkbookManager wm, int chapNum) {
    Sheet sheet = wm.wb.getSheet("BookNames");
    Row row = sheet.getRow(chapNum);
    String bookName = row.getCell(1).getStringCellValue();
    String chapterTitle = row.getCell(2).getStringCellValue();
    return chapterTitle.replace("_", bookName);
  }

  /** Given one verse, publish it as called for in the spreadsheet
   * @param bkno 2-digit book number
   * @param line verse with chapter abbreviation, space, chapter:verse 2 spaces,
   * then the verse text.
   * @param wm workbook manager
   * @param doc Word .docx file being generated
   * @return modified line of text
   */
  public static String publishVerse(String bkno, String line, WorkbookManager wm, XWPFDocument doc) {
    // Parse the verse reference (e.g., "LUK 1:1")
    if (line.length() < 7) {
      return null;
    }

    // Extract chapter and verse numbers and text
    String[] parts = line.substring(4).split(":", 2);
    if (parts.length < 2) {
      return null;
    }

    String chapterNum = parts[0].trim();

    // Split verse number from text
    String remaining = parts[1];
    int spaceIndex = remaining.indexOf(' ');
    if (spaceIndex == -1) {
      return null;
    }

    String verseNum = remaining.substring(0, spaceIndex);
    String verseText = remaining.substring(spaceIndex + 2); // Skip the spaces after verse number

    // If verse 1, add chapter heading and verse with drop cap
    if ("1".equals(verseNum)) {
      // Add chapter heading
      XWPFParagraph chapterPara = doc.createParagraph();
      chapterPara.setAlignment(ParagraphAlignment.CENTER);

      // Set "keep with next" paragraph attribute
      CTPPr chapterPPr = chapterPara.getCTP().isSetPPr() ? chapterPara.getCTP().getPPr() : chapterPara.getCTP().addNewPPr();
      chapterPPr.addNewKeepNext();

      XWPFRun run = chapterPara.createRun();
      run.setBold(true);
      run.setText("Chapter " + chapterNum);

      // Add verse with superscript verse number  // TODO drop cap
      if (verseText.length() > 0) {
	XWPFParagraph versePara = doc.createParagraph();
	versePara.setStyle("FAH");

	// Add superscript verse number
	XWPFRun verseNumRun = versePara.createRun();
	verseNumRun.setText(verseNum);
	verseNumRun.setSubscript(VerticalAlign.SUPERSCRIPT);

	// Add verse text
	XWPFRun textRun = versePara.createRun();
	textRun.setText(verseText);
      }
    } else {
      // Add verse with superscript verse number
      if (verseText.length() > 0) {
	XWPFParagraph versePara = doc.createParagraph();
	versePara.setStyle("FAH");

	// Add superscript verse number
	XWPFRun verseNumRun = versePara.createRun();
	verseNumRun.setText(verseNum);
	verseNumRun.setSubscript(VerticalAlign.SUPERSCRIPT);

	// Add verse text
	XWPFRun textRun = versePara.createRun();
	textRun.setText(verseText);
      }
    }
    return null;
  }
}
