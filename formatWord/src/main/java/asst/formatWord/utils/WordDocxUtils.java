package asst.formatWord.utils;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import java.io.FileOutputStream;
import java.math.BigInteger;

/** A set of magic spells to add new items to a XWPFDocument doc.</p>
 * 
 * <p>Footnotes and bookmarks are numbered starting at one in the document.
 * The footnote and bookmark counters are not thread-safe.
 * When the document is inserted into another Word document, Word
 * renumbers bookmarks and footnotes to fit with footnotes and
 * bookmarks which are already there.  If a bookmark name conflicts
 * with a bookmark which is already in the document, which one is kept doesn't
 * matter - either way, the results won't be what you expect.</p>
 * <p>The unit test program shows how to poke around in the Word .xml
 * model in great detail.  That structure is sufficiently arcane that
 * the methods in this file for dealing with them can rightfully
 * be regarded as magic spells.
 *
 * @author GitHub CoPilot
 * @since 2025 12
 */
public class WordDocxUtils {

  private static int footnoteCounter = 1;  // not thread safe
  private static int bookmarkCounter = 1;  // not thread safe

  /** Superscript spell */
  public static XWPFParagraph addSuperscriptParagraph(XWPFDocument doc,
      String superText,
      String normalText) {
    XWPFParagraph para = doc.createParagraph();
    XWPFRun superRun = para.createRun();
    superRun.setText(superText);
    superRun.setSubscript(VerticalAlign.SUPERSCRIPT);

    XWPFRun normalRun = para.createRun();
    normalRun.setText(normalText);
    return para;
  }

  /** Footnote spells */
  public static void addFootnote(XWPFParagraph para, XWPFDocument doc, String footnoteText) {
    XWPFRun run = para.createRun();
    run.setStyle("Footnote Reference");
    CTFtnEdnRef ref = run.getCTR().addNewFootnoteReference();
    ref.setId(BigInteger.valueOf(footnoteCounter));

    XWPFFootnote footnote = doc.createFootnote();
    footnote.getCTFtnEdn().setId(BigInteger.valueOf(footnoteCounter));
    footnote.createParagraph().createRun().setText(footnoteText);

    footnoteCounter++;
  }

  /** Add a paragraph with a footnote reference at a specified position in the text 
   * @param para The paragraph to add the footnote to
   * @param doc The document (needed to create the footnote)
   * @param text The paragraph text
   * @param where Position in the text where the footnote reference should appear (0 = beginning, text.length() or greater = end)
   * @param footnoteText The text of the footnote
   */
  public static void addFootnote(XWPFParagraph para, XWPFDocument doc, String text, int where,
      String footnoteText) {
    // Ensure where is within valid bounds
    if (where < 0) {
      where = 0;
    }
    if (where > text.length()) {
      where = text.length();
    }

    // If footnote goes at the beginning
    if (where == 0) {
      // Create run with footnote reference
      XWPFRun footnoteRun = para.createRun();
      footnoteRun.setStyle("Footnote Reference");
      CTFtnEdnRef ref = footnoteRun.getCTR().addNewFootnoteReference();
      ref.setId(BigInteger.valueOf(footnoteCounter));

      XWPFFootnote footnote = doc.createFootnote();
      footnote.getCTFtnEdn().setId(BigInteger.valueOf(footnoteCounter));
      footnote.createParagraph().createRun().setText(footnoteText);
      footnoteCounter++;

      // Add the text after the footnote
      if (text.length() > 0) {
        XWPFRun textRun = para.createRun();
        textRun.setText(text);
      }
    }
    // If footnote goes at the end
    else if (where >= text.length()) {
      // Add the text first
      XWPFRun textRun = para.createRun();
      textRun.setText(text);

      // Then add footnote reference
      XWPFRun footnoteRun = para.createRun();
      footnoteRun.setStyle("Footnote Reference");
      CTFtnEdnRef ref = footnoteRun.getCTR().addNewFootnoteReference();
      ref.setId(BigInteger.valueOf(footnoteCounter));

      XWPFFootnote footnote = doc.createFootnote();
      footnote.getCTFtnEdn().setId(BigInteger.valueOf(footnoteCounter));
      footnote.createParagraph().createRun().setText(footnoteText);
      footnoteCounter++;
    }
    // Footnote goes in the middle
    else {
      // Add text before the footnote
      String beforeText = text.substring(0, where);
      XWPFRun beforeRun = para.createRun();
      beforeRun.setText(beforeText);

      // Add footnote reference
      XWPFRun footnoteRun = para.createRun();
      footnoteRun.setStyle("Footnote Reference");
      CTFtnEdnRef ref = footnoteRun.getCTR().addNewFootnoteReference();
      ref.setId(BigInteger.valueOf(footnoteCounter));

      XWPFFootnote footnote = doc.createFootnote();
      footnote.getCTFtnEdn().setId(BigInteger.valueOf(footnoteCounter));
      footnote.createParagraph().createRun().setText(footnoteText);
      footnoteCounter++;

      // Add text after the footnote
      String afterText = text.substring(where);
      XWPFRun afterRun = para.createRun();
      afterRun.setText(afterText);
    }
  }

  /** Bookmark spell */
  public static void addBookmarkParagraph(XWPFDocument doc,
      String bookmarkName,
      String text) {
    XWPFParagraph para = doc.createParagraph();
    XWPFRun run = para.createRun();
    run.setText(text);

    CTBookmark bookmarkStart = para.getCTP().addNewBookmarkStart();
    bookmarkStart.setId(BigInteger.valueOf(bookmarkCounter));
    bookmarkStart.setName(bookmarkName);

    para.getCTP().addNewBookmarkEnd().setId(BigInteger.valueOf(bookmarkCounter));
    bookmarkCounter++;
  }

  /** Index entry spell */
  public static void addIndexEntry(XWPFParagraph para, String entryText) {
    XWPFRun runBegin = para.createRun();
    CTFldChar fldCharBegin = runBegin.getCTR().addNewFldChar();
    fldCharBegin.setFldCharType(STFldCharType.BEGIN);

    XWPFRun runInstr = para.createRun();
    runInstr.getCTR().addNewInstrText().setStringValue(" XE \"" + entryText + "\" ");

    XWPFRun runSep = para.createRun();
    CTFldChar fldCharSep = runSep.getCTR().addNewFldChar();
    fldCharSep.setFldCharType(STFldCharType.SEPARATE);

    XWPFRun runEnd = para.createRun();
    CTFldChar fldCharEnd = runEnd.getCTR().addNewFldChar();
    fldCharEnd.setFldCharType(STFldCharType.END);
  }

  /** Split heading spell */
  public static void addSplitHeading(XWPFDocument doc,
      String headingText,
      String trailingText) {
    XWPFParagraph headingPara = doc.createParagraph();
    headingPara.setStyle("Heading2");
    headingPara.createRun().setText(headingText);

    XWPFParagraph normalPara = doc.createParagraph();
    normalPara.setStyle("Normal");
    normalPara.setSpacingBefore(0);
    normalPara.setSpacingAfter(0);
    normalPara.createRun().setText(trailingText);
  }

  /** Create a single paragraph with Heading2 style where only the first part
   * appears in the TOC, but the second part is visible in the document.
   * This is achieved by creating two paragraphs where the first has a hidden
   * paragraph marker (vanish and specVanish).
   * @param doc The document
   * @param headingText Text that will appear in TOC
   * @param trailingText Additional text that won't appear in TOC
   */
  public static void addSplitHeading2Para(XWPFDocument doc,
      String headingText,
      String trailingText) {
    // First paragraph with Heading2 style
    XWPFParagraph headingPara = doc.createParagraph();
    headingPara.setStyle("Heading2");

    // Add the heading text
    XWPFRun headingRun = headingPara.createRun();
    headingRun.setText(headingText);

    if ((trailingText == null) || trailingText.isEmpty()) { return; }

    // Make the paragraph marker hidden (this is the style separator)
    CTPPr pPr = headingPara.getCTP().isSetPPr() ? headingPara.getCTP().getPPr() : headingPara.getCTP().addNewPPr();
    org.openxmlformats.schemas.wordprocessingml.x2006.main.CTParaRPr rPr = pPr.addNewRPr();
    rPr.addNewVanish();
    rPr.addNewSpecVanish();

    // Second paragraph with the trailing text (no heading style)
    XWPFParagraph trailingPara = doc.createParagraph();
    XWPFRun trailingRun = trailingPara.createRun();
    trailingRun.setText(" " + trailingText);
  }

  /**
   * Adds a hyperlink run in the given paragraph that points to a bookmark.
   * @param para        the paragraph to add the hyperlink into
   * @param bookmarkName the name of the bookmark to jump to
   * @param linkText     the visible text of the hyperlink
   */
  public static void addHyperlinkToBookmark(XWPFParagraph para,
      String bookmarkName,
      String linkText) {
    // Create the hyperlink element
    CTHyperlink ctHyperlink = para.getCTP().addNewHyperlink();
    ctHyperlink.setAnchor(bookmarkName); // anchor = bookmark name

    // Create a run with hyperlink styling
    CTR ctr = CTR.Factory.newInstance();
    CTRPr rpr = ctr.addNewRPr();
    rpr.addNewColor().setVal("0000FF"); // blue
    rpr.addNewU().setVal(STUnderline.SINGLE); // underline

    ctr.addNewT().setStringValue(linkText);

    // Attach run to hyperlink
    ctHyperlink.addNewR().set(ctr);
  }

  /**
   * Sets the document to update all fields (including table of contents) when opened in Word.
   * This is necessary because Apache POI cannot directly update field codes.
   * @param doc the document to configure
   */
  public static void setUpdateFieldsOnOpen(XWPFDocument doc) {
    try {
      org.apache.xmlbeans.XmlCursor cursor = doc.getDocument().newCursor();
      cursor.toFirstChild();
      cursor.beginElement("settings", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
      cursor.beginElement("updateFields", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
      cursor.insertAttributeWithValue("val", "true");
      cursor.dispose();
    } catch (Exception e) {
      System.err.println("Warning: Could not set updateFields property: " + e.getMessage());
    }
  }

  // Example save method
  public static void saveDoc(XWPFDocument doc, String filename) throws Exception {
    try (FileOutputStream out = new FileOutputStream(filename)) {
      doc.write(out);
    }
  }

}
