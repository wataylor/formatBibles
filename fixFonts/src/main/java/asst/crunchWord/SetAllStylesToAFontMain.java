package asst.crunchWord;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBaseStyles;
import org.openxmlformats.schemas.drawingml.x2006.main.CTFontCollection;
import org.openxmlformats.schemas.drawingml.x2006.main.CTFontScheme;
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

  /**
   * @param args will be a .docx file eventually.
   * @throws Exception
   */
  public static void main(String[] args) throws Exception {
    String inputPath;

    if (args.length < 2) {
      System.out.println("This program requires two command-line arguments:\n"
	  + "The name of the desired font\n"
	  + "The path to the .docx file to have all fonts set.\n"
	  + "Be sure to enclose the font or path in double quotes if they have spaces.\n"
	  + "\nWARNING:  NO error checking.  If you select an unknown font, the .docx may become unreadable.");
      System.exit(1);
    }

    fontToEmbed = args[0];
    inputPath = args[1];

    System.out.println("Setting all fonts to: " + fontToEmbed);
    System.out.println("Processing: " + inputPath);

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
    System.out.println("Number of styles (getStyleList): " + (ctStyles.getStyleList() != null ? ctStyles.getStyleList().size() : 0));
    System.out.println("Number of styles (sizeOfStyleArray): " + ctStyles.sizeOfStyleArray());
    System.out.println("Has document defaults: " + ctStyles.isSetDocDefaults());

    // DEBUG: Output first 1000 characters of the XML to see what's actually in there
    String xmlContent = ctStyles.xmlText();
    System.out.println("\n=== First 1000 chars of styles.xml ===");
    System.out.println(xmlContent.substring(0, Math.min(1000, xmlContent.length())));
    System.out.println("=== End of sample ===\n");

    // Count actual <w:style> elements in the XML
    int styleTagCount = xmlContent.split("<w:style ").length - 1;
    System.out.println("Number of <w:style> tags found in XML: " + styleTagCount);

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
	System.out.println("  Updated style: " + styleName + " (" + styleType + ") -> " + fontToEmbed);
      }
    }
    System.out.println("Processed " + styleCount + " styles (" + paragraphStyleCount + " paragraph, " + characterStyleCount + " character)");

    // Write updated styles.xml back into the package
    try (OutputStream os = stylesPart.getOutputStream()) {
      stylesDoc.save(os);
    }

    // Locate /word/numbering.xml
    PackagePart numberingPart = null;
    for (PackagePart part : pkg.getParts()) {
      if (part.getPartName().getName().equals("/word/numbering.xml")) {
	numberingPart = part;
	break;
      }
    }

    if (numberingPart != null) {
      // Parse numbering.xml
      CTNumbering ctNumbering;
      try (InputStream is = numberingPart.getInputStream()) {
	ctNumbering = CTNumbering.Factory.parse(is);
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
	  fonts.setAscii(fontToEmbed);
	  fonts.setHAnsi(fontToEmbed);
	  fonts.setCs(fontToEmbed);
	  fonts.setEastAsia(fontToEmbed);
	}
      }

      // Write updated numbering.xml back
      try (OutputStream os = numberingPart.getOutputStream()) {
	ctNumbering.save(os);
      }
    }

    // Locate /word/theme/theme1.xml
    PackagePart themePart = null;
    for (PackagePart part : pkg.getParts()) {
      if (part.getPartName().getName().equals("/word/theme/theme1.xml")) {
	themePart = part;
	break;
      }
    }

    if (themePart != null) {

      // Parse theme1.xml
      ThemeDocument themeDoc;
      try (InputStream is = themePart.getInputStream()) {
	themeDoc = ThemeDocument.Factory.parse(is);
      }

      // Access the theme
      org.openxmlformats.schemas.drawingml.x2006.main.CTOfficeStyleSheet theme = themeDoc.getTheme();
      CTBaseStyles themeElements = theme.getThemeElements();
      CTFontScheme fontScheme = themeElements.getFontScheme();

      // Major and minor font collections
      CTFontCollection major = fontScheme.getMajorFont();
      CTFontCollection minor = fontScheme.getMinorFont();

      // Major Latin
      if (major.getLatin() != null) {
	major.getLatin().setTypeface(fontToEmbed);
      }

      // Minor Latin
      if (minor.getLatin() != null) {
	minor.getLatin().setTypeface(fontToEmbed);
      }

      // East Asian fonts
      if (major.getEa() != null) {
	major.getEa().setTypeface(fontToEmbed);
      }
      if (minor.getEa() != null) {
	minor.getEa().setTypeface(fontToEmbed);
      }

      // Complex script fonts
      if (major.getCs() != null) {
	major.getCs().setTypeface(fontToEmbed);
      }
      if (minor.getCs() != null) {
	minor.getCs().setTypeface(fontToEmbed);
      }

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

      // Save updated theme1.xml
      try (OutputStream os = themePart.getOutputStream()) {
	themeDoc.save(os);
      }
    }

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

    // Close package (writes changes)  Word cares not how it was changed.
    pkg.close();

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

    fonts.setAscii(fontToEmbed);
    fonts.setHAnsi(fontToEmbed);
    fonts.setCs(fontToEmbed);
    fonts.setEastAsia(fontToEmbed);
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

	  try (OutputStream os = part.getOutputStream()) {
	    ftrDoc.save(os);
	  }
	  System.out.println("Processed " + partName);
	}
      }
    }
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

      try (OutputStream os = commentPart.getOutputStream()) {
	commentsDoc.save(os);
      }
      System.out.println("Processed comments.xml");
    }
  }

  /** Process fontTable.xml - clear all fonts and add only Times New Roman */
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
      if (fontsList != null) {
	// Remove ALL existing font declarations
	while (fontsList.sizeOfFontArray() > 0) {
	  fontsList.removeFont(0);
	}

	// Add only Times New Roman
	CTFont timesFont = fontsList.addNewFont();
	timesFont.setName(fontToEmbed);

	System.out.println("Cleared fontTable.xml and added only: " + fontToEmbed);
      }

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

