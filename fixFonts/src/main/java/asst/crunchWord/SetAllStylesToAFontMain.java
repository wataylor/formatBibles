package asst.crunchWord;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.ByteArrayOutputStream;
import java.net.URL;
import java.net.URLConnection;
import java.net.URI;
import java.nio.charset.StandardCharsets;
import java.nio.file.FileSystems;
import java.nio.file.FileSystem;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Arrays;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import javax.xml.namespace.QName;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBaseStyles;
import org.openxmlformats.schemas.drawingml.x2006.main.CTFontCollection;
import org.openxmlformats.schemas.drawingml.x2006.main.CTFontScheme;
import org.openxmlformats.schemas.drawingml.x2006.main.CTOfficeStyleSheet;
import java.io.ByteArrayInputStream;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSupplementalFont;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextFont;
import org.openxmlformats.schemas.drawingml.x2006.main.ThemeDocument;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTComment;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTComments;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFont;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFontsList;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFtnEdn;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHdrFtr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTNumbering;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import java.io.FileOutputStream;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPrGeneral;import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSettings;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyles;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CommentsDocument;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.DocumentDocument;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.EndnotesDocument;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.FontsDocument;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.FootnotesDocument;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.FtrDocument;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.HdrDocument;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.NumberingDocument;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.SettingsDocument;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.StylesDocument;
import org.apache.xmlbeans.XmlOptions;

/** Read a .docx file and set all the fonts to Calibri. Word tends
 * to hide font references in many places in a document.  This results
 * in creating .pdf files which refer to fonts which are not used but
 * are not embedded because Word only embeds visible fonts.  This produces
 * errors when trying to use a .pdf in an app that requires embedded fonts.

 * All styles
 * Document defaults
 * Theme fonts (including all script-specific fallbacks)
 * All content (document, headers, footers, tables, hyperlinks, footnotes, endnotes)
 * Font table

 * @author Material Gain
 * @since 2025 12 
 */
public class SetAllStylesToAFontMain {

  /** The font to use for all text in the document */
  public static String fontToEmbed = "Calibri";

  /** Enable extra diagnostic logging when true. */
  public static boolean verbose = false;

  /** Current input document path for debug output context. */
  public static String currentInputPath = null;

  /**
   * @param args will be a .docx file eventually.
   * @throws Exception
   */
  public static void main(String[] args) throws Exception {
    String inputPath = null;

    if (args.length < 2) {
      printUsage();
      System.exit(1);
    }

    int positionalArgCount = 0;
    for (String arg : args) {
      if ("-verbose".equalsIgnoreCase(arg)) {
	verbose = true;
	continue;
      }

      if (positionalArgCount == 0) {
	fontToEmbed = arg;
      } else if (positionalArgCount == 1) {
	inputPath = arg;
      } else {
	printUsage();
	System.exit(1);
      }
      positionalArgCount++;
    }

    if (positionalArgCount != 2 || inputPath == null) {
      printUsage();
      System.exit(1);
    }

    currentInputPath = inputPath;

    printRuntimeBuildInfo();
    System.out.println("Setting all fonts to: " + fontToEmbed);
    System.out.println("Processing: " + inputPath);
    logVerbose("Verbose logging enabled.");

    // Open the DOCX as a low-level OPC package
    OPCPackage pkg = OPCPackage.open(new File(inputPath), PackageAccess.READ_WRITE);

    // Locate /word/styles.xml
    PackagePart stylesPart = null;
    for (PackagePart part : pkg.getParts()) {
      if (part.getPartName().getName().equals("/word/styles.xml")) {
	stylesPart = part;
	break;
      }
    }

    if (stylesPart == null) {
      System.out.println("No styles.xml found in " + inputPath);
      pkg.close();
      return;
    }

    System.out.println("Found styles.xml, parsing...");

    // Parse styles.xml using StylesDocument (proper wrapper)
    StylesDocument stylesDoc;
    try (InputStream is = stylesPart.getInputStream()) {
      stylesDoc = StylesDocument.Factory.parse(is);
    }

    CTStyles ctStyles = stylesDoc.getStyles();

    System.out.println("Parsed styles.xml successfully");
    logVerbose("Number of styles (getStyleList): " + (ctStyles.getStyleList() != null ? ctStyles.getStyleList().size() : 0));
    logVerbose("Number of styles (sizeOfStyleArray): " + ctStyles.sizeOfStyleArray());
    logVerbose("Has document defaults: " + ctStyles.isSetDocDefaults());

    if (verbose) {
      String xmlContent = ctStyles.xmlText();
      logVerbose("\n=== First 1000 chars of styles.xml ===");
      logVerbose(xmlContent.substring(0, Math.min(1000, xmlContent.length())));
      logVerbose("=== End of sample ===\n");

      int styleTagCount = xmlContent.split("<w:style ").length - 1;
      logVerbose("Number of <w:style> tags found in XML: " + styleTagCount);
    }

    // Process document defaults first (these affect all styles that don't override fonts)
    if (ctStyles.isSetDocDefaults()) {
      org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocDefaults docDefaults = ctStyles.getDocDefaults();

      // Set default run properties
      if (docDefaults.isSetRPrDefault()) {
	org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPrDefault rprDefault = docDefaults.getRPrDefault();
	CTRPr rpr = rprDefault.isSetRPr() ? rprDefault.getRPr() : rprDefault.addNewRPr();

	CTFonts fonts = null;
	if (rpr.sizeOfRFontsArray() > 0) {
	  fonts = rpr.getRFontsArray(0);
	} else {
	  fonts = rpr.addNewRFonts();
	}

	fonts.setAscii(fontToEmbed);
	fonts.setHAnsi(fontToEmbed);
	fonts.setCs(fontToEmbed);
	fonts.setEastAsia(fontToEmbed);
      }
    }

    // Iterate all styles and force fonts to specified font
    int styleCount = 0;
    int characterStyleCount = 0;
    int paragraphStyleCount = 0;
    // Use array-based access (POI 5.2.4 pattern)
    for (int i = 0; i < ctStyles.sizeOfStyleArray(); i++) {
      CTStyle ctStyle = ctStyles.getStyleArray(i);

      // ALWAYS add run properties to every style, regardless of current state
      CTRPr rpr = ctStyle.isSetRPr() ? ctStyle.getRPr() : ctStyle.addNewRPr();
      CTFonts fonts = null;
      if (rpr.sizeOfRFontsArray() > 0) {
	fonts = rpr.getRFontsArray(0);
      } else {
	fonts = rpr.addNewRFonts();
      }

      String styleName = ctStyle.isSetName() ? ctStyle.getName().getVal() : ctStyle.getStyleId();
      String styleType = ctStyle.isSetType() ? ctStyle.getType().toString() : "unknown";

      fonts.setAscii(fontToEmbed);
      fonts.setHAnsi(fontToEmbed);
      fonts.setCs(fontToEmbed);
      fonts.setEastAsia(fontToEmbed);
      styleCount++;

      if ("character".equals(styleType)) {
	characterStyleCount++;
      } else if ("paragraph".equals(styleType)) {
	paragraphStyleCount++;
      }

      // Log key styles to verify they're being processed
      if (styleName != null && (styleName.equals("Normal") || styleName.equals("FAH") || 
	  styleName.contains("TOC") || styleName.contains("HTML") || 
	  styleName.contains("Hyperlink") || styleName.contains("Link"))) {
	logVerbose("  Updated style: " + styleName + " (" + styleType + ") -> " + fontToEmbed);
      }
    }
    System.out.println("Processed " + styleCount + " styles (" + paragraphStyleCount + " paragraph, " + characterStyleCount + " character)");

    // Force invisible defaults as well (docDefaults + latentStyles)
    processLatentAndDefaultFonts(stylesDoc, "/word/styles.xml");

    // Write updated styles.xml back into the package
    try (OutputStream os = stylesPart.getOutputStream()) {
      stylesDoc.save(os);
    }

    // Process optional /word/stylesWithEffects.xml if present
    processStylesWithEffectsXml(pkg);

    // Locate /word/numbering.xml
    PackagePart numberingPart = null;
    for (PackagePart part : pkg.getParts()) {
      if (part.getPartName().getName().equals("/word/numbering.xml")) {
	numberingPart = part;
	break;
      }
    }

    byte[] finalNumberingBytes = null;

    if (numberingPart != null) {
      // Parse numbering.xml using wrapper document to preserve proper root serialization
      NumberingDocument numberingDoc;
      try (InputStream is = numberingPart.getInputStream()) {
    	numberingDoc = NumberingDocument.Factory.parse(is);
      }

      CTNumbering ctNumbering = numberingDoc.getNumbering();
      if (ctNumbering == null) {
    	ctNumbering = numberingDoc.addNewNumbering();
      }

      // Iterate all abstract numbering definitions
      for (CTAbstractNum abs : ctNumbering.getAbstractNumList()) {
	for (CTLvl lvl : abs.getLvlList()) {

	  // Ensure run properties exist
	  CTRPr rpr = lvl.isSetRPr() ? lvl.getRPr() : lvl.addNewRPr();

	  // Ensure fonts exist (POI 5.2.4 uses arrays)
	  CTFonts fonts = null;
	  if (rpr.sizeOfRFontsArray() > 0) {
	    fonts = rpr.getRFontsArray(0);
	  } else {
	    fonts = rpr.addNewRFonts();
	  }

    // Force all fonts to specified font
    setAllFontFamilies(fonts);
	}
      }

      // Also rewrite all existing <w:rFonts> in numbering.xml (e.g., lvlOverride entries)
      XmlObject[] numberingFonts = ctNumbering.selectPath(
    "declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:rFonts");
      for (XmlObject fontObj : numberingFonts) {
  if (fontObj instanceof CTFonts) {
    setAllFontFamilies((CTFonts) fontObj);
  }
      }
      logVerbose("Updated " + numberingFonts.length + " numbering.xml <w:rFonts> elements");

      // Write updated numbering.xml back atomically to avoid any append-style stream behavior
      byte[] numberingBytes;
      try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
    	numberingDoc.save(bos);
    	numberingBytes = normalizeNumberingXmlBytes(bos.toByteArray());
      }
      finalNumberingBytes = numberingBytes;

      File debugNumbering = writeNumberingDebugFile(numberingBytes);
      System.out.println("numbering.xml debug file: " + debugNumbering.getAbsolutePath());

      PackagePartName numberingPartName = numberingPart.getPartName();
      String numberingContentType = numberingPart.getContentType();
      pkg.removePart(numberingPartName);
      PackagePart newNumberingPart = pkg.createPart(numberingPartName, numberingContentType);
      try (OutputStream os = newNumberingPart.getOutputStream()) {
    	os.write(numberingBytes);
      }
    }

        // Process all theme XML parts (usually only theme1.xml)
        processThemeParts(pkg);

        // Process DrawingML font declarations in chart/diagram/drawing parts
        processDrawingMlParts(pkg);

    // Process document.xml (main document content)
    processDocumentXml(pkg);

    // Process headers (header1.xml, header2.xml, header3.xml)
    processPartsByPattern(pkg, "/word/header", "Header");

    // Process footers (footer1.xml, footer2.xml, footer3.xml)
    processPartsByPattern(pkg, "/word/footer", "Footer");

    // Process footnotes.xml
    processFootnotesXml(pkg);

    // Process endnotes.xml
    processEndnotesXml(pkg);

    // Process comments.xml
    processCommentsXml(pkg);

    // Process fontTable.xml
    processFontTableXml(pkg);

    // Process settings.xml
    processSettingsXml(pkg);

    // Final fallback sweep across all /word/*.xml parts for any remaining w:rFonts
    processAllWordPartsForRFonts(pkg);

    // Final alias sweep: replace lingering Arial-like font names with the requested font
    replaceFontAliasesInAllWordParts(pkg);

    // Diagnostic: report any remaining Arial-like tokens anywhere in the package
    reportAliasTokensInPackage(pkg);

    // Close package (writes changes)  Word cares not how it was changed.
    pkg.close();

    // Final hardening for numbering.xml: write directly into ZIP with truncate semantics
    if (finalNumberingBytes != null) {
      forceWriteNumberingXmlInZip(inputPath, finalNumberingBytes);
      validateNumberingXmlInDocx(inputPath);
    }

    System.out.println("\n========================================");
    System.out.println("All fonts set to " + fontToEmbed + " in all document parts.");
    System.out.println("Document saved: " + inputPath);
    System.out.println("========================================");

  }

  /** Helper method to set fonts in run properties */
  private static void setFontsInRPr(CTRPr rpr) {
    CTFonts fonts = null;
    if (rpr.sizeOfRFontsArray() > 0) {
      fonts = rpr.getRFontsArray(0);
    } else {
      fonts = rpr.addNewRFonts();
    }

    setAllFontFamilies(fonts);
  }

  /** Process /word/stylesWithEffects.xml if present. */
  private static void processStylesWithEffectsXml(OPCPackage pkg) throws Exception {
    PackagePart stylesWithEffectsPart = null;
    for (PackagePart part : pkg.getParts()) {
      if (part.getPartName().getName().equals("/word/stylesWithEffects.xml")) {
	stylesWithEffectsPart = part;
	break;
      }
    }

    if (stylesWithEffectsPart == null) {
      logVerbose("No stylesWithEffects.xml found");
      return;
    }

    StylesDocument stylesDoc;
    try (InputStream is = stylesWithEffectsPart.getInputStream()) {
      stylesDoc = StylesDocument.Factory.parse(is);
    }

    CTStyles ctStyles = stylesDoc.getStyles();
    if (ctStyles == null) {
      logVerbose("stylesWithEffects.xml has no <w:styles> root");
      return;
    }

    if (ctStyles.isSetDocDefaults()) {
      org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocDefaults docDefaults = ctStyles.getDocDefaults();
      if (docDefaults.isSetRPrDefault()) {
	org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPrDefault rprDefault = docDefaults.getRPrDefault();
	CTRPr rpr = rprDefault.isSetRPr() ? rprDefault.getRPr() : rprDefault.addNewRPr();
	setFontsInRPr(rpr);
      }
    }

    int updatedStyles = 0;
    for (int i = 0; i < ctStyles.sizeOfStyleArray(); i++) {
      CTStyle ctStyle = ctStyles.getStyleArray(i);
      CTRPr rpr = ctStyle.isSetRPr() ? ctStyle.getRPr() : ctStyle.addNewRPr();
      setFontsInRPr(rpr);
      updatedStyles++;
    }

    // Force invisible defaults as well (docDefaults + latentStyles)
    processLatentAndDefaultFonts(stylesDoc, "/word/stylesWithEffects.xml");

    try (OutputStream os = stylesWithEffectsPart.getOutputStream()) {
      stylesDoc.save(os);
    }

    System.out.println("Processed stylesWithEffects.xml (updated " + updatedStyles + " styles)");
  }

  /** Set all key WordprocessingML font families to the desired font. */
  private static void setAllFontFamilies(CTFonts fonts) {
    fonts.setAscii(fontToEmbed);
    fonts.setHAnsi(fontToEmbed);
    fonts.setCs(fontToEmbed);
    fonts.setEastAsia(fontToEmbed);

    // Clear theme-driven font mappings so Word cannot resolve to unexpected fonts (e.g., Hyperlink -> Arial)
    if (fonts.isSetAsciiTheme()) {
      fonts.unsetAsciiTheme();
    }
    if (fonts.isSetHAnsiTheme()) {
      fonts.unsetHAnsiTheme();
    }
    if (fonts.isSetEastAsiaTheme()) {
      fonts.unsetEastAsiaTheme();
    }
    if (fonts.isSetCstheme()) {
      fonts.unsetCstheme();
    }
  }

  /**
   * Rewrite latent/default style font declarations that Word may apply implicitly.
   */
  private static void processLatentAndDefaultFonts(StylesDocument stylesDoc, String partName) {
    XmlObject[] fontNodes = stylesDoc.selectPath(
	"declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' "
	    + ".//w:docDefaults//w:rFonts | .//w:latentStyles//w:rFonts");

    for (XmlObject node : fontNodes) {
      if (node instanceof CTFonts) {
	setAllFontFamilies((CTFonts) node);
      }

      try (XmlCursor cursor = node.newCursor()) {
	cursor.setAttributeText(new QName("", "ascii"), fontToEmbed);
	cursor.setAttributeText(new QName("", "hAnsi"), fontToEmbed);
	cursor.setAttributeText(new QName("", "cs"), fontToEmbed);
	cursor.setAttributeText(new QName("", "eastAsia"), fontToEmbed);

	cursor.removeAttribute(new QName("", "asciiTheme"));
	cursor.removeAttribute(new QName("", "hAnsiTheme"));
	cursor.removeAttribute(new QName("", "eastAsiaTheme"));
	cursor.removeAttribute(new QName("", "cstheme"));
	cursor.removeAttribute(new QName("", "csTheme"));
      }
    }

    if (fontNodes.length > 0) {
      logVerbose("Updated " + fontNodes.length + " latent/default <w:rFonts> nodes in " + partName);
    } else {
      logVerbose("No latent/default <w:rFonts> nodes found in " + partName);
    }
  }

  /** Print extra logging only when -verbose is enabled. */
  private static void logVerbose(String message) {
    if (verbose) {
      System.out.println(message);
    }
  }

  /** Print command-line usage. */
  private static void printUsage() {
    System.out.println("Usage: SetAllStylesToAFontMain [-verbose] <fontName> <docxPath>\n"
	+ "  -verbose : optional extra diagnostics (single dash)\n"
	+ "  <fontName> : desired font\n"
	+ "  <docxPath> : path to the .docx file\n"
	+ "Be sure to enclose font or path in double quotes if they have spaces.\n"
	+ "\nWARNING: NO error checking. If you select an unknown font, the .docx may become unreadable.");
  }

  /** Print runtime build identity so it's obvious which compiled class is running. */
  private static void printRuntimeBuildInfo() {
    try {
      String classPath = SetAllStylesToAFontMain.class.getName().replace('.', '/') + ".class";
      URL classUrl = SetAllStylesToAFontMain.class.getClassLoader().getResource(classPath);
      String classLocation = classUrl != null ? classUrl.toString() : "unknown";

      String compiledAt = "unknown";
      if (classUrl != null) {
	URLConnection connection = classUrl.openConnection();
	long modified = connection.getLastModified();
	if (modified > 0) {
	  compiledAt = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS Z").format(new Date(modified));
	}
      }

      String implVersion = SetAllStylesToAFontMain.class.getPackage() != null
	  ? SetAllStylesToAFontMain.class.getPackage().getImplementationVersion()
	  : null;

      System.out.println("Runtime class: " + classLocation);
      System.out.println("Runtime class timestamp: " + compiledAt);
      System.out.println("Runtime implementation version: " + (implVersion != null ? implVersion : "(not set)"));
    } catch (Exception e) {
      System.out.println("Runtime build info unavailable: " + e.getMessage());
    }
  }

  /** Ensure a theme font collection has all key entries set to the desired font. */
  private static void setThemeCollectionFonts(CTFontCollection collection) {
    CTTextFont latin = collection.getLatin() != null ? collection.getLatin() : collection.addNewLatin();
    latin.setTypeface(fontToEmbed);

    CTTextFont eastAsia = collection.getEa() != null ? collection.getEa() : collection.addNewEa();
    eastAsia.setTypeface(fontToEmbed);

    CTTextFont complexScript = collection.getCs() != null ? collection.getCs() : collection.addNewCs();
    complexScript.setTypeface(fontToEmbed);
  }

  /** Process all theme XML parts under /word/theme/*.xml */
  private static void processThemeParts(OPCPackage pkg) throws Exception {
    int processedThemes = 0;

    for (PackagePart themePart : pkg.getParts()) {
      String partName = themePart.getPartName().getName();
      if (!partName.startsWith("/word/theme/") || !partName.endsWith(".xml")) {
	continue;
      }

      ThemeDocument themeDoc;
      try (InputStream is = themePart.getInputStream()) {
	themeDoc = ThemeDocument.Factory.parse(is);
      }

      CTOfficeStyleSheet theme = themeDoc.getTheme();
      CTBaseStyles themeElements = theme.getThemeElements();
      CTFontScheme fontScheme = themeElements.getFontScheme();

      // Major and minor font collections
      CTFontCollection major = fontScheme.getMajorFont();
      CTFontCollection minor = fontScheme.getMinorFont();

      // Ensure major/minor latin, east-asia, and complex-script fonts exist and are set
      setThemeCollectionFonts(major);
      setThemeCollectionFonts(minor);

      // Update ALL script-specific fonts (Arab, Hebr, Thai, Viet, etc.)
      if (major.sizeOfFontArray() > 0) {
	for (CTSupplementalFont scriptFont : major.getFontList()) {
	  scriptFont.setTypeface(fontToEmbed);
	}
      }
      if (minor.sizeOfFontArray() > 0) {
	for (CTSupplementalFont scriptFont : minor.getFontList()) {
	  scriptFont.setTypeface(fontToEmbed);
	}
      }

      try (OutputStream os = themePart.getOutputStream()) {
	themeDoc.save(os);
      }
      processedThemes++;
      System.out.println("Processed " + partName + " (major/minor latin, ea, cs, and supplemental script fonts)");
    }

    if (processedThemes == 0) {
      System.out.println("No theme XML parts found under /word/theme/");
    }
  }

  /** Process all runs in a list of paragraphs */
  private static void processParagraphs(List<CTP> paragraphs) {
    for (CTP para : paragraphs) {
      // Process ALL runs in the paragraph
      if (para.sizeOfRArray() > 0) {
	for (CTR run : para.getRArray()) {
	  // ALWAYS set run properties, even if not already present
	  CTRPr rpr = run.isSetRPr() ? run.getRPr() : run.addNewRPr();
	  setFontsInRPr(rpr);

	  // Process drawing objects (shapes, text boxes, etc.) within runs
	  if (run.sizeOfDrawingArray() > 0) {
	    processDrawingObjects(run);
	  }
	}
      } else {
	// Paragraph has no runs (empty paragraph) - add a dummy run with font properties
	CTR newRun = para.addNewR();
	CTRPr rpr = newRun.addNewRPr();
	setFontsInRPr(rpr);
      }

      // ALSO process hyperlink runs (hyperlinks contain their own runs)
      if (para.sizeOfHyperlinkArray() > 0) {
	for (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink hyperlink : para.getHyperlinkList()) {
	  if (hyperlink.sizeOfRArray() > 0) {
	    for (CTR run : hyperlink.getRArray()) {
	      CTRPr rpr = run.isSetRPr() ? run.getRPr() : run.addNewRPr();
	      setFontsInRPr(rpr);
	    }
	  }
	}
      }
    }
  }

  /** Process DrawingML objects (shapes, text boxes) within a run */
  private static void processDrawingObjects(CTR run) {
    try {
      // Use XmlObject to process drawing elements without needing specific drawing classes
      if (run.sizeOfDrawingArray() > 0) {
	// Access drawing elements as generic XmlObjects
	for (int i = 0; i < run.sizeOfDrawingArray(); i++) {
	  XmlObject drawing = run.getDrawingArray(i);
	  processDrawingMLFonts(drawing);
	}
      }
    } catch (Exception e) {
      // DrawingML processing can be complex, log but don't fail
      System.out.println("Warning: Could not fully process drawing object fonts: " + e.getMessage());
    }
  }

  /** Recursively process fonts in DrawingML objects */
  private static void processDrawingMLFonts(XmlObject obj) {
    try {
      // Use XPath to find all text properties in the drawing object
      // This handles shapes, text boxes, and other drawing elements
      String xml = obj.xmlText();

      // Look for latin, ea, cs font references in the DrawingML namespace
      if (xml.contains("<a:latin") || xml.contains("<a:ea") || xml.contains("<a:cs")) {
	// Parse and update fonts in the drawing object
	XmlObject[] textFonts = obj.selectPath(
	    "declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' " +
		".//a:latin | .//a:ea | .//a:cs"
	    );

	for (XmlObject fontObj : textFonts) {
	  if (fontObj instanceof CTTextFont) {
	    CTTextFont textFont = (CTTextFont) fontObj;
	    textFont.setTypeface(fontToEmbed);
	  }
	}
      }
    } catch (Exception e) {
      // DrawingML structure can vary, continue processing
      System.out.println("Warning: Could not process some DrawingML fonts: " + e.getMessage());
    }
  }

    /**
     * Process chart/diagram/drawing XML parts and force all DrawingML latin/ea/cs fonts
     * to use the requested typeface.
     */
    private static void processDrawingMlParts(OPCPackage pkg) throws Exception {
      int partsProcessed = 0;
      int fontNodesUpdated = 0;

      for (PackagePart part : pkg.getParts()) {
    String partName = part.getPartName().getName();
    if (!partName.startsWith("/word/") || !partName.endsWith(".xml")) {
  	continue;
    }

    if (isHandledByDedicatedProcessor(partName)) {
  	continue;
    }

    boolean isDrawingMlPart = partName.startsWith("/word/charts/")
  	  || partName.startsWith("/word/diagrams/")
  	  || partName.startsWith("/word/drawings/");
    if (!isDrawingMlPart) {
  	continue;
    }

    XmlObject xmlDoc;
    try (InputStream is = part.getInputStream()) {
  	xmlDoc = XmlObject.Factory.parse(is);
    }

    XmlObject[] textFonts = xmlDoc.selectPath(
  	  "declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' "
  	      + ".//a:latin | .//a:ea | .//a:cs");

    if (textFonts.length == 0) {
  	continue;
    }

    for (XmlObject fontObj : textFonts) {
  	try (XmlCursor cursor = fontObj.newCursor()) {
  	  cursor.setAttributeText(new QName("", "typeface"), fontToEmbed);
  	}
  	fontNodesUpdated++;
    }

    try (OutputStream os = part.getOutputStream()) {
  	xmlDoc.save(os);
    }

    partsProcessed++;
    logVerbose("Processed DrawingML fonts in " + partName + " (updated " + textFonts.length + " nodes)");
      }

      if (partsProcessed > 0) {
    System.out.println("Processed DrawingML font declarations in " + partsProcessed
  	  + " part(s), updated " + fontNodesUpdated + " node(s)");
      } else {
    logVerbose("No chart/diagram/drawing parts with DrawingML font declarations found");
      }
    }

  /** Process document.xml */
  private static void processDocumentXml(OPCPackage pkg) throws Exception {
    PackagePart docPart = null;
    for (PackagePart part : pkg.getParts()) {
      if (part.getPartName().getName().equals("/word/document.xml")) {
	docPart = part;
	break;
      }
    }

    if (docPart != null) {
      DocumentDocument docDoc;
      try (InputStream is = docPart.getInputStream()) {
	docDoc = DocumentDocument.Factory.parse(is);
      }

      CTDocument1 document = docDoc.getDocument();
      CTBody body = document.getBody();

      // Process all paragraphs in the document
      if (body.sizeOfPArray() > 0) {
	processParagraphs(body.getPList());
      }

      // Process all tables in the document
      if (body.sizeOfTblArray() > 0) {
	for (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl table : body.getTblList()) {
	  processTable(table);
	}
      }

      // Safety sweep: force fonts on every run in document.xml, including runs
      // that may be outside the normal paragraph/table traversal paths.
      int runsUpdated = processAllRunsInXml(document, "/word/document.xml");
      logVerbose("Document-wide run sweep updated " + runsUpdated + " runs");

      // Save updated document.xml
      try (OutputStream os = docPart.getOutputStream()) {
	docDoc.save(os);
      }
      System.out.println("Processed document.xml");
    }
  }

  /** Process all text in a table */
  private static void processTable(org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl table) {
    // Process each row
    if (table.sizeOfTrArray() > 0) {
      for (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow row : table.getTrList()) {
	// Process each cell
	if (row.sizeOfTcArray() > 0) {
	  for (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc cell : row.getTcList()) {
	    // Process paragraphs in the cell
	    if (cell.sizeOfPArray() > 0) {
	      processParagraphs(cell.getPList());
	    }
	    // Process nested tables
	    if (cell.sizeOfTblArray() > 0) {
	      for (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl nestedTable : cell.getTblList()) {
		processTable(nestedTable);
	      }
	    }
	  }
	}
      }
    }
  }

  /**
   * Force font settings on every WordprocessingML run in a part-level XML object.
   * Returns number of runs touched.
   */
  private static int processAllRunsInXml(XmlObject root, String partName) {
    XmlObject[] runs = root.selectPath(
	"declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:r");

    int updated = 0;
    for (XmlObject runObj : runs) {
      if (runObj instanceof CTR) {
	CTR run = (CTR) runObj;
	CTRPr rpr = run.isSetRPr() ? run.getRPr() : run.addNewRPr();
	setFontsInRPr(rpr);
	updated++;
      }
    }

    logVerbose("Run sweep in " + partName + " found " + runs.length + " runs, updated " + updated);
    return updated;
  }

      /**
       * Final safety pass: normalize all w:rFonts nodes in every /word/*.xml part.
       */
      private static void processAllWordPartsForRFonts(OPCPackage pkg) throws Exception {
        int partsProcessed = 0;
        int fontNodesUpdated = 0;
        int skippedMalformedParts = 0;

        for (PackagePart part : pkg.getParts()) {
          String partName = part.getPartName().getName();
          if (!partName.startsWith("/word/") || !partName.endsWith(".xml")) {
  continue;
          }
          if (isHandledByDedicatedProcessor(partName)) {
  continue;
          }

          byte[] partBytes;
          try (InputStream is = part.getInputStream(); ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
  byte[] buffer = new byte[8192];
  int read;
  while ((read = is.read(buffer)) >= 0) {
    bos.write(buffer, 0, read);
  }
  partBytes = bos.toByteArray();
          }

          XmlObject xmlDoc = null;
          try {
  xmlDoc = XmlObject.Factory.parse(new ByteArrayInputStream(partBytes));
          } catch (Exception firstParseEx) {
  byte[] repairedBytes = normalizeXmlDeclarations(partBytes);
  if (!Arrays.equals(partBytes, repairedBytes)) {
    try {
      xmlDoc = XmlObject.Factory.parse(new ByteArrayInputStream(repairedBytes));
      try (OutputStream os = part.getOutputStream()) {
        os.write(repairedBytes);
      }
      logVerbose("Recovered malformed XML declarations in " + partName + " during fallback sweep");
    } catch (Exception secondParseEx) {
      File dump = writeDebugFile(repairedBytes, "fallback-malformed-" + partName.replace('/', '_'));
      System.out.println("Warning: Skipping malformed part in fallback sweep: " + partName
    + " (debug: " + dump.getAbsolutePath() + ")");
      skippedMalformedParts++;
      continue;
    }
  } else {
    File dump = writeDebugFile(partBytes, "fallback-malformed-" + partName.replace('/', '_'));
    System.out.println("Warning: Skipping malformed part in fallback sweep: " + partName
        + " (debug: " + dump.getAbsolutePath() + ")");
    skippedMalformedParts++;
    continue;
  }
          }

          XmlObject[] fontNodes = xmlDoc.selectPath(
        "declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:rFonts");

          if (fontNodes.length == 0) {
  continue;
          }

          for (XmlObject fontNode : fontNodes) {
  if (fontNode instanceof CTFonts) {
    setAllFontFamilies((CTFonts) fontNode);
  }

  try (XmlCursor cursor = fontNode.newCursor()) {
    cursor.setAttributeText(new QName("", "ascii"), fontToEmbed);
    cursor.setAttributeText(new QName("", "hAnsi"), fontToEmbed);
    cursor.setAttributeText(new QName("", "cs"), fontToEmbed);
    cursor.setAttributeText(new QName("", "eastAsia"), fontToEmbed);

    cursor.removeAttribute(new QName("", "asciiTheme"));
    cursor.removeAttribute(new QName("", "hAnsiTheme"));
    cursor.removeAttribute(new QName("", "eastAsiaTheme"));
    cursor.removeAttribute(new QName("", "cstheme"));
    cursor.removeAttribute(new QName("", "csTheme"));
  }
  fontNodesUpdated++;
          }

          try (OutputStream os = part.getOutputStream()) {
  xmlDoc.save(os);
          }

  partsProcessed++;
  logVerbose("Fallback rFonts sweep processed " + partName + " (" + fontNodes.length + " nodes)");
        }

        System.out.println("Fallback rFonts sweep updated " + fontNodesUpdated + " nodes across " + partsProcessed + " /word/*.xml parts");
        if (skippedMalformedParts > 0) {
  System.out.println("Fallback rFonts sweep skipped malformed parts: " + skippedMalformedParts);
        }
      }

      /** True if this XML part already has dedicated processing logic elsewhere. */
      private static boolean isHandledByDedicatedProcessor(String partName) {
        if (partName.equals("/word/document.xml")
      || partName.equals("/word/styles.xml")
      || partName.equals("/word/stylesWithEffects.xml")
      || partName.equals("/word/numbering.xml")
      || partName.equals("/word/fontTable.xml")
      || partName.equals("/word/settings.xml")
      || partName.equals("/word/footnotes.xml")
      || partName.equals("/word/endnotes.xml")
      || partName.equals("/word/comments.xml")) {
          return true;
        }

        if (partName.startsWith("/word/header")
      || partName.startsWith("/word/footer")
      || partName.startsWith("/word/theme/")
      || partName.startsWith("/word/charts/")
      || partName.startsWith("/word/diagrams/")
      || partName.startsWith("/word/drawings/")) {
          return true;
        }

        return false;
      }

      /**
       * Normalize numbering.xml bytes to a single XML declaration and a single numbering root.
       */
      private static byte[] normalizeNumberingXmlBytes(byte[] xmlBytes) {
        String xml = new String(xmlBytes, StandardCharsets.UTF_8);

        int rootStart = xml.indexOf("<w:numbering");
        if (rootStart < 0) {
          rootStart = xml.indexOf("<numbering");
        }

        int rootEndStart = xml.lastIndexOf("</w:numbering>");
        int rootEndLen = "</w:numbering>".length();
        if (rootEndStart < 0) {
          rootEndStart = xml.lastIndexOf("</numbering>");
          rootEndLen = "</numbering>".length();
        }

        if (rootStart < 0 || rootEndStart < 0 || rootEndStart <= rootStart) {
          return xmlBytes;
        }

        String rootXml = xml.substring(rootStart, rootEndStart + rootEndLen);
        rootXml = rootXml.replaceAll("<\\?xml[^>]*\\?>", "").trim();

        String normalized = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" + rootXml;
        return normalized.getBytes(StandardCharsets.UTF_8);
      }

      /** Write normalized numbering.xml bytes to a debug file and return canonical file path. */
      private static File writeNumberingDebugFile(byte[] bytes) {
        return writeDebugFile(bytes, "numbering-final-debug.xml");
      }

      /** Write debug bytes alongside the input .docx and return canonical file path. */
      private static File writeDebugFile(byte[] bytes, String fileName) {
        try {
          File inputFile = currentInputPath != null ? new File(currentInputPath).getAbsoluteFile() : null;
          File baseDir = inputFile != null ? inputFile.getParentFile() : new File(".").getAbsoluteFile();
          File outFile = new File(baseDir, fileName).getCanonicalFile();

          try (FileOutputStream fos = new FileOutputStream(outFile, false)) {
    	fos.write(bytes);
          }
          return outFile;
        } catch (Exception e) {
          throw new RuntimeException("Could not write numbering debug file", e);
        }
      }

      /**
       * Validate numbering.xml as stored in the final DOCX zip (exact bytes Word will read).
       */
      private static void validateNumberingXmlInDocx(String docxPath) {
        try (ZipFile zip = new ZipFile(docxPath)) {
          ZipEntry entry = zip.getEntry("word/numbering.xml");
          if (entry == null) {
    	logVerbose("No word/numbering.xml entry found during post-write validation");
    	return;
          }

          byte[] bytes;
          try (InputStream is = zip.getInputStream(entry); ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
    	byte[] buf = new byte[8192];
    	int read;
    	while ((read = is.read(buf)) >= 0) {
    	  bos.write(buf, 0, read);
    	}
    	bytes = bos.toByteArray();
          }

          int declCount = countXmlDeclarationOccurrences(bytes);
          logVerbose("Readback numbering.xml XML declaration count: " + declCount);

          try (ByteArrayInputStream bis = new ByteArrayInputStream(bytes)) {
    	NumberingDocument.Factory.parse(bis);
          } catch (Exception parseEx) {
      // Attempt automatic repair by re-normalizing declarations and rewriting once
      byte[] repaired = normalizeXmlDeclarations(bytes);
      if (!Arrays.equals(bytes, repaired)) {
        forceWriteNumberingXmlInZip(docxPath, repaired);
        try (ByteArrayInputStream repairedBis = new ByteArrayInputStream(repaired)) {
          NumberingDocument.Factory.parse(repairedBis);
        }
        logVerbose("Repaired numbering.xml declarations after readback parse failure");
        return;
      }

      File readbackDump = writeDebugFile(bytes, "numbering-readback-debug.xml");
      System.out.println("numbering.xml readback debug file: " + readbackDump.getAbsolutePath());
      throw new RuntimeException("numbering.xml readback parse failed. Debug file: "
          + readbackDump.getAbsolutePath(), parseEx);
          }

          if (declCount != 1) {
    	File readbackDump = writeDebugFile(bytes, "numbering-readback-debug.xml");
      System.out.println("numbering.xml readback debug file: " + readbackDump.getAbsolutePath());
    	throw new RuntimeException("numbering.xml readback declaration count is " + declCount
    	    + ". Debug file: " + readbackDump.getAbsolutePath());
          }
        } catch (RuntimeException re) {
          throw re;
        } catch (Exception e) {
          throw new RuntimeException("Post-write numbering.xml validation failed", e);
        }
      }

  /** Process header or footer files by pattern */
  private static void processPartsByPattern(OPCPackage pkg, String pattern, String type) throws Exception {
    for (PackagePart part : pkg.getParts()) {
      String partName = part.getPartName().getName();
      if (partName.startsWith(pattern) && partName.endsWith(".xml")) {
	if (type.equals("Header")) {
	  HdrDocument hdrDoc;
	  try (InputStream is = part.getInputStream()) {
	    hdrDoc = HdrDocument.Factory.parse(is);
	  }

	  CTHdrFtr hdr = hdrDoc.getHdr();
	  if (hdr.sizeOfPArray() > 0) {
	    processParagraphs(hdr.getPList());
	  }
	  if (hdr.sizeOfTblArray() > 0) {
	    for (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl table : hdr.getTblList()) {
	      processTable(table);
	    }
	  }

    int runsUpdated = processAllRunsInXml(hdr, partName);
    logVerbose("Header run sweep updated " + runsUpdated + " runs in " + partName);

	  try (OutputStream os = part.getOutputStream()) {
	    hdrDoc.save(os);
	  }
	  System.out.println("Processed " + partName);
	} else if (type.equals("Footer")) {
	  FtrDocument ftrDoc;
	  try (InputStream is = part.getInputStream()) {
	    ftrDoc = FtrDocument.Factory.parse(is);
	  }

	  CTHdrFtr ftr = ftrDoc.getFtr();
	  if (ftr.sizeOfPArray() > 0) {
	    processParagraphs(ftr.getPList());
	  }
	  if (ftr.sizeOfTblArray() > 0) {
	    for (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl table : ftr.getTblList()) {
	      processTable(table);
	    }
	  }

    int runsUpdated = processAllRunsInXml(ftr, partName);
    logVerbose("Footer run sweep updated " + runsUpdated + " runs in " + partName);

	  try (OutputStream os = part.getOutputStream()) {
	    ftrDoc.save(os);
	  }
	  System.out.println("Processed " + partName);
	}
      }
    }
  }

  /**
   * Remove duplicate/embedded XML declarations from every /word/*.xml part before parsing.
   */
  private static void sanitizeWordXmlDeclarations(OPCPackage pkg) throws Exception {
    int sanitizedParts = 0;
    for (PackagePart part : pkg.getParts()) {
      String partName = part.getPartName().getName();
      if (!partName.startsWith("/word/") || !partName.endsWith(".xml")) {
	continue;
      }

      byte[] original;
      try (InputStream is = part.getInputStream(); ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
	byte[] buf = new byte[8192];
	int read;
	while ((read = is.read(buf)) >= 0) {
	  bos.write(buf, 0, read);
	}
	original = bos.toByteArray();
      }

      byte[] normalized = normalizeXmlDeclarations(original);
      if (!Arrays.equals(original, normalized)) {
	try (OutputStream os = part.getOutputStream()) {
	  os.write(normalized);
	}
	sanitizedParts++;
	logVerbose("Sanitized XML declarations in " + partName);
      }
    }

    if (sanitizedParts > 0) {
      System.out.println("Sanitized XML declarations in " + sanitizedParts + " /word/*.xml part(s)");
    }
  }

  /**
   * Keep at most one XML declaration at the top and remove any embedded declarations.
   */
  private static byte[] normalizeXmlDeclarations(byte[] bytes) {
    String xml = new String(bytes, StandardCharsets.UTF_8);
    String firstDecl = null;

    String trimmedLeading = xml.replaceFirst("^\\uFEFF", "");
    java.util.regex.Matcher startDecl = java.util.regex.Pattern
	.compile("(?is)^\\s*<\\?xml[^>]*\\?>")
	.matcher(trimmedLeading);
    if (startDecl.find()) {
      firstDecl = startDecl.group().trim();
    }

    String withoutDecls = trimmedLeading.replaceAll("(?is)<\\?xml[^>]*\\?>", "").trim();
    String normalized = (firstDecl != null ? firstDecl + "\n" : "") + withoutDecls;
    return normalized.getBytes(StandardCharsets.UTF_8);
  }

  /**
   * Overwrite /word/numbering.xml directly in the .docx zip to avoid append behavior.
   */
  private static void forceWriteNumberingXmlInZip(String docxPath, byte[] numberingBytes) {
    Map<String, String> env = new HashMap<>();
    env.put("create", "false");

    URI zipUri = URI.create("jar:" + new File(docxPath).toURI().toString());
    try (FileSystem zipFs = FileSystems.newFileSystem(zipUri, env)) {
      Path numberingPath = zipFs.getPath("/word/numbering.xml");
      Files.write(numberingPath, numberingBytes,
	  StandardOpenOption.CREATE,
	  StandardOpenOption.TRUNCATE_EXISTING,
	  StandardOpenOption.WRITE);

      byte[] verify = Files.readAllBytes(numberingPath);
      int declCount = countOccurrences(verify, "<?xml");
      logVerbose("Post-ZIP numbering.xml XML declaration count: " + declCount);
    } catch (Exception e) {
      throw new RuntimeException("Could not force-write /word/numbering.xml in docx zip", e);
    }
  }

  /** Count simple substring occurrences in UTF-8 byte content. */
  private static int countOccurrences(byte[] bytes, String needle) {
    String text = new String(bytes, StandardCharsets.UTF_8);
    int count = 0;
    int fromIndex = 0;
    while (true) {
      int idx = text.indexOf(needle, fromIndex);
      if (idx < 0) {
	break;
      }
      count++;
      fromIndex = idx + needle.length();
    }
    return count;
  }

  /** Count XML declarations case-insensitively. */
  private static int countXmlDeclarationOccurrences(byte[] bytes) {
    String text = new String(bytes, StandardCharsets.UTF_8);
    java.util.regex.Matcher matcher = java.util.regex.Pattern
	.compile("(?is)<\\?xml\\b")
	.matcher(text);
    int count = 0;
    while (matcher.find()) {
      count++;
    }
    return count;
  }

  /**
   * Replace lingering font aliases (e.g., Arial) in /word/*.xml parts.
   */
  private static void replaceFontAliasesInAllWordParts(OPCPackage pkg) throws Exception {
    String[] aliases = new String[] { "Arial", "ArialMT", "Ariel", "Aries" };

    int scannedParts = 0;
    int partsTouched = 0;
    int replacements = 0;
    int skippedMalformedParts = 0;

    for (PackagePart part : pkg.getParts()) {
      String partName = part.getPartName().getName();
      if (!partName.startsWith("/word/") || !partName.endsWith(".xml")) {
	continue;
      }
      scannedParts++;

      String xmlText;
      try (InputStream is = part.getInputStream(); ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
    	byte[] buffer = new byte[8192];
    	int read;
    	while ((read = is.read(buffer)) >= 0) {
    	  bos.write(buffer, 0, read);
    	}
    	xmlText = new String(bos.toByteArray(), StandardCharsets.UTF_8);
      } catch (Exception ex) {
  skippedMalformedParts++;
  logVerbose("Alias sweep skipped malformed part: " + partName + " -> " + ex.getMessage());
  continue;
      }

      int[] replacementCount = new int[] { 0 };
      String updatedXml = replaceAliasValuesInXmlText(xmlText, aliases, fontToEmbed, replacementCount);
      int partReplacements = replacementCount[0];
      if (partReplacements == 0) {
	continue;
      }

      try (OutputStream os = part.getOutputStream()) {
    	os.write(updatedXml.getBytes(StandardCharsets.UTF_8));
      }

      partsTouched++;
      replacements += partReplacements;
      logVerbose("Alias sweep replaced " + partReplacements + " values in " + partName);
    }

    System.out.println("Alias sweep scanned " + scannedParts + " /word/*.xml part(s), replaced "
	+ replacements + " Arial-like values across " + partsTouched + " part(s)"
	+ (skippedMalformedParts > 0 ? "; skipped malformed: " + skippedMalformedParts : ""));
  }

  /**
   * Diagnostic pass: find raw Arial-like tokens in any package part bytes.
   */
  private static void reportAliasTokensInPackage(OPCPackage pkg) throws Exception {
    String[] needles = new String[] { "arial", "ariel", "aries" };
    int partsWithHits = 0;
    int skippedUnreadable = 0;

    for (PackagePart part : pkg.getParts()) {
      String partName = part.getPartName().getName();
      byte[] bytes;
      try (InputStream is = part.getInputStream(); ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
  byte[] buffer = new byte[8192];
  int read;
  while ((read = is.read(buffer)) >= 0) {
    bos.write(buffer, 0, read);
  }
  bytes = bos.toByteArray();
      } catch (InvalidOperationException ioe) {
  skippedUnreadable++;
  logVerbose("Alias token diagnostic skipped unreadable part: " + partName + " -> " + ioe.getMessage());
  continue;
      } catch (Exception ex) {
  skippedUnreadable++;
  logVerbose("Alias token diagnostic skipped part due to read error: " + partName + " -> " + ex.getMessage());
  continue;
      }

      String text = new String(bytes, StandardCharsets.UTF_8).toLowerCase();
      int hits = 0;
      for (String needle : needles) {
	int idx = 0;
	while (true) {
	  idx = text.indexOf(needle, idx);
	  if (idx < 0) {
	    break;
	  }
	  hits++;
	  idx += needle.length();
	}
      }

      if (hits > 0) {
	partsWithHits++;
	System.out.println("Alias token diagnostic: " + partName + " -> " + hits + " hit(s)");
      }
    }

    if (partsWithHits == 0) {
      System.out.println("Alias token diagnostic: no Arial-like tokens found in package parts");
    }
    if (skippedUnreadable > 0) {
      logVerbose("Alias token diagnostic skipped unreadable parts: " + skippedUnreadable);
    }
  }

  /** Replace font aliases in key font-related XML attributes. */
  private static String replaceAliasValuesInXmlText(String xml, String[] aliases, String targetFont, int[] replacementCount) {
    String aliasPattern = buildAliasRegexAlternation(aliases);

    // Pass 1: known font-related attribute names
    Pattern pattern = Pattern.compile(
	"(?i)(\\b(?:ascii|hAnsi|cs|eastAsia|asciiTheme|hAnsiTheme|eastAsiaTheme|cstheme|csTheme|typeface|name|val)\\s*=\\s*\"|\\b(?:ascii|hAnsi|cs|eastAsia|asciiTheme|hAnsiTheme|eastAsiaTheme|cstheme|csTheme|typeface|name|val)\\s*=\\s*')"
	    + "([^\"']*(?:" + aliasPattern + ")[^\"']*)"
	    + "(\"|')");

    Matcher matcher = pattern.matcher(xml);
    StringBuffer sb = new StringBuffer();
    int count = 0;
    while (matcher.find()) {
      matcher.appendReplacement(sb,
	  Matcher.quoteReplacement(matcher.group(1) + targetFont + matcher.group(3)));
      count++;
    }
    matcher.appendTail(sb);

    // Pass 2: any attribute value containing an alias token
    String pass1Xml = sb.toString();
    Pattern anyAttrPattern = Pattern.compile(
	"(?i)(\\b[\\w:.-]+\\s*=\\s*\"|\\b[\\w:.-]+\\s*=\\s*')"
	    + "([^\"']*(?:" + aliasPattern + ")[^\"']*)"
	    + "(\"|')");

    Matcher anyAttrMatcher = anyAttrPattern.matcher(pass1Xml);
    StringBuffer sb2 = new StringBuffer();
    while (anyAttrMatcher.find()) {
      anyAttrMatcher.appendReplacement(sb2,
	  Matcher.quoteReplacement(anyAttrMatcher.group(1) + targetFont + anyAttrMatcher.group(3)));
      count++;
    }
    anyAttrMatcher.appendTail(sb2);

    replacementCount[0] = count;
    return sb2.toString();
  }

  /** Build safe regex alternation for alias strings. */
  private static String buildAliasRegexAlternation(String[] aliases) {
    StringBuilder sb = new StringBuilder();
    for (int i = 0; i < aliases.length; i++) {
      if (i > 0) {
	sb.append("|");
      }
      sb.append(Pattern.quote(aliases[i]));
    }
    return sb.toString();
  }

  /** Process footnotes.xml */
  private static void processFootnotesXml(OPCPackage pkg) throws Exception {
    PackagePart footnotePart = null;
    for (PackagePart part : pkg.getParts()) {
      if (part.getPartName().getName().equals("/word/footnotes.xml")) {
	footnotePart = part;
	break;
      }
    }

    if (footnotePart != null) {
      FootnotesDocument fnDoc;
      try (InputStream is = footnotePart.getInputStream()) {
	fnDoc = FootnotesDocument.Factory.parse(is);
      }

      // Process each footnote
      if (fnDoc.getFootnotes().sizeOfFootnoteArray() > 0) {
	for (CTFtnEdn footnote : fnDoc.getFootnotes().getFootnoteList()) {
	  if (footnote.sizeOfPArray() > 0) {
	    processParagraphs(footnote.getPList());
	  }
	}
      }

      int runsUpdated = processAllRunsInXml(fnDoc.getFootnotes(), "/word/footnotes.xml");
      logVerbose("Footnotes run sweep updated " + runsUpdated + " runs");

      try (OutputStream os = footnotePart.getOutputStream()) {
	fnDoc.save(os);
      }
      System.out.println("Processed footnotes.xml");
    }
  }

  /** Process endnotes.xml */
  private static void processEndnotesXml(OPCPackage pkg) throws Exception {
    PackagePart endnotePart = null;
    for (PackagePart part : pkg.getParts()) {
      if (part.getPartName().getName().equals("/word/endnotes.xml")) {
	endnotePart = part;
	break;
      }
    }

    if (endnotePart != null) {
      EndnotesDocument enDoc;
      try (InputStream is = endnotePart.getInputStream()) {
	enDoc = EndnotesDocument.Factory.parse(is);
      }

      // Process each endnote
      if (enDoc.getEndnotes().sizeOfEndnoteArray() > 0) {
	for (CTFtnEdn endnote : enDoc.getEndnotes().getEndnoteList()) {
	  if (endnote.sizeOfPArray() > 0) {
	    processParagraphs(endnote.getPList());
	  }
	}
      }

      int runsUpdated = processAllRunsInXml(enDoc.getEndnotes(), "/word/endnotes.xml");
      logVerbose("Endnotes run sweep updated " + runsUpdated + " runs");

      try (OutputStream os = endnotePart.getOutputStream()) {
	enDoc.save(os);
      }
      System.out.println("Processed endnotes.xml");
    }
  }

  /** Process comments.xml */
  private static void processCommentsXml(OPCPackage pkg) throws Exception {
    PackagePart commentPart = null;
    for (PackagePart part : pkg.getParts()) {
      if (part.getPartName().getName().equals("/word/comments.xml")) {
	commentPart = part;
	break;
      }
    }

    if (commentPart != null) {
      CommentsDocument commentsDoc;
      try (InputStream is = commentPart.getInputStream()) {
	commentsDoc = CommentsDocument.Factory.parse(is);
      }

      CTComments comments = commentsDoc.getComments();
      if (comments != null && comments.sizeOfCommentArray() > 0) {
	for (CTComment comment : comments.getCommentList()) {
	  if (comment.sizeOfPArray() > 0) {
	    processParagraphs(comment.getPList());
	  }
	}
      }

      if (comments != null) {
  int runsUpdated = processAllRunsInXml(comments, "/word/comments.xml");
  logVerbose("Comments run sweep updated " + runsUpdated + " runs");
      }

      try (OutputStream os = commentPart.getOutputStream()) {
	commentsDoc.save(os);
      }
      System.out.println("Processed comments.xml");
    }
  }

  /** Process fontTable.xml - clear all fonts and add only the requested font */
  private static void processFontTableXml(OPCPackage pkg) throws Exception {
    PackagePart fontTablePart = null;
    for (PackagePart part : pkg.getParts()) {
      if (part.getPartName().getName().equals("/word/fontTable.xml")) {
	fontTablePart = part;
	break;
      }
    }

    if (fontTablePart != null) {
      FontsDocument fontsDoc;
      try (InputStream is = fontTablePart.getInputStream()) {
	fontsDoc = FontsDocument.Factory.parse(is);
      }

      CTFontsList fontsList = fontsDoc.getFonts();
      if (fontsList == null) {
  fontsList = fontsDoc.addNewFonts();
      }

  // Remove ALL existing font declarations
  while (fontsList.sizeOfFontArray() > 0) {
    fontsList.removeFont(0);
  }

  // Add only the requested font
  CTFont targetFont = fontsList.addNewFont();
  targetFont.setName(fontToEmbed);

  System.out.println("Cleared fontTable.xml and added only: " + fontToEmbed);

      try (OutputStream os = fontTablePart.getOutputStream()) {
	fontsDoc.save(os);
      }
    }
  }

  /** Process settings.xml - set default fonts */
  private static void processSettingsXml(OPCPackage pkg) throws Exception {
    PackagePart settingsPart = null;
    for (PackagePart part : pkg.getParts()) {
      if (part.getPartName().getName().equals("/word/settings.xml")) {
	settingsPart = part;
	break;
      }
    }

    if (settingsPart != null) {
      SettingsDocument settingsDoc;
      try (InputStream is = settingsPart.getInputStream()) {
	settingsDoc = SettingsDocument.Factory.parse(is);
      }

      CTSettings settings = settingsDoc.getSettings();

      // Rewrite document-level default fonts if present
      XmlObject[] defaultFontsNodes = settings.selectPath(
    	  "declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:defaultFonts");
      for (XmlObject defaultFontsNode : defaultFontsNodes) {
    	try (XmlCursor cursor = defaultFontsNode.newCursor()) {
    	  cursor.setAttributeText(new QName("", "ascii"), fontToEmbed);
    	  cursor.setAttributeText(new QName("", "hAnsi"), fontToEmbed);
    	  cursor.setAttributeText(new QName("", "cs"), fontToEmbed);
    	  cursor.setAttributeText(new QName("", "eastAsia"), fontToEmbed);

    // Remove theme-based mappings so they cannot override explicit font selection
    cursor.removeAttribute(new QName("", "asciiTheme"));
    cursor.removeAttribute(new QName("", "hAnsiTheme"));
    cursor.removeAttribute(new QName("", "eastAsiaTheme"));
    cursor.removeAttribute(new QName("", "cstheme"));
    cursor.removeAttribute(new QName("", "csTheme"));
    	}
      }
      if (defaultFontsNodes.length > 0) {
    	System.out.println("Updated settings.xml defaultFonts entries: " + defaultFontsNodes.length);
      } else {
    	logVerbose("No <w:defaultFonts> entries found in settings.xml");
      }

      // Set theme font mappings if they exist
      if (settings.isSetThemeFontLang()) {
	// Theme fonts are already handled in theme1.xml processing
      }

      // Process compatibility settings fonts if present
      if (settings.isSetCompat()) {
	// Compatibility settings may reference fonts but typically don't define them directly
      }

      try (OutputStream os = settingsPart.getOutputStream()) {
	settingsDoc.save(os);
      }
      System.out.println("Processed settings.xml");
    }
  }
}

