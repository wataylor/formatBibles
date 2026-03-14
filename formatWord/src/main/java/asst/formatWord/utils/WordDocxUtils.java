package asst.formatWord.utils;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.apache.xmlbeans.XmlCursor;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

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

  /** Superscript spell
   * @param doc The Word document
   * @param superText Text to be in superscript
   * @param normalText Text to be normal
   * @return A new paragraph */
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

  /** Footnote spells
   * @param para Paragraph in the document
   * @param doc The Word document
   * @param footnoteText The text of the footnote */
  public static void addFootnote(XWPFParagraph para, XWPFDocument doc, String footnoteText) {
    XWPFRun run = para.createRun();
    run.setStyle("FootnoteReference");
    CTFtnEdnRef ref = run.getCTR().addNewFootnoteReference();
    ref.setId(BigInteger.valueOf(footnoteCounter));

    XWPFFootnote footnote = doc.createFootnote();
    footnote.getCTFtnEdn().setId(BigInteger.valueOf(footnoteCounter));
    XWPFParagraph footnotePara = footnote.createParagraph();
    footnotePara.setStyle("FootnoteText");
    footnotePara.createRun().setText(footnoteText);
    applyItalicTags(footnotePara, doc);

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
      footnoteRun.setStyle("FootnoteReference");
      CTFtnEdnRef ref = footnoteRun.getCTR().addNewFootnoteReference();
      ref.setId(BigInteger.valueOf(footnoteCounter));

      XWPFFootnote footnote = doc.createFootnote();
      footnote.getCTFtnEdn().setId(BigInteger.valueOf(footnoteCounter));
      XWPFParagraph footnotePara = footnote.createParagraph();
      footnotePara.setStyle("FootnoteText");
      footnotePara.createRun().setText(footnoteText);
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
      footnoteRun.setStyle("FootnoteReference");
      CTFtnEdnRef ref = footnoteRun.getCTR().addNewFootnoteReference();
      ref.setId(BigInteger.valueOf(footnoteCounter));

      XWPFFootnote footnote = doc.createFootnote();
      footnote.getCTFtnEdn().setId(BigInteger.valueOf(footnoteCounter));
      XWPFParagraph footnotePara = footnote.createParagraph();
      footnotePara.setStyle("FootnoteText");
      footnotePara.createRun().setText(footnoteText);
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
      footnoteRun.setStyle("FootnoteReference");
      CTFtnEdnRef ref = footnoteRun.getCTR().addNewFootnoteReference();
      ref.setId(BigInteger.valueOf(footnoteCounter));

      XWPFFootnote footnote = doc.createFootnote();
      footnote.getCTFtnEdn().setId(BigInteger.valueOf(footnoteCounter));
      XWPFParagraph footnotePara = footnote.createParagraph();
      footnotePara.setStyle("FootnoteText");
      footnotePara.createRun().setText(footnoteText);
      footnoteCounter++;

      // Add text after the footnote
      String afterText = text.substring(where);
      XWPFRun afterRun = para.createRun();
      afterRun.setText(afterText);
    }
  }

  /**
   * Replace balanced HTML-like italic tags in paragraph text with true italic runs.
   * The markers &lt;i&gt; and &lt;/i&gt; are removed from visible text.
   *
   * <p>Example: {@code this is <i>very</i> important} becomes three runs:
   * normal "this is ", italic "very", normal " important".</p>
   *
   * @param para paragraph whose text may contain balanced {@code <i>} markers
   * @param doc document containing the paragraph (included for API consistency)
   */
  public static void applyItalicTags(XWPFParagraph para, XWPFDocument doc) {
    List<XWPFRun> originalRuns = para.getRuns();
    if (originalRuns == null || originalRuns.isEmpty()) {
      return;
    }

    StringBuilder fullTextBuilder = new StringBuilder();
    List<XWPFRun> runByCharIndex = new ArrayList<>();
    for (XWPFRun run : originalRuns) {
      String runText = run.text();
      if (runText == null || runText.isEmpty()) {
        continue;
      }
      for (int i = 0; i < runText.length(); i++) {
        fullTextBuilder.append(runText.charAt(i));
        runByCharIndex.add(run);
      }
    }

    String text = fullTextBuilder.toString();
    if (text.isEmpty()) {
      return;
    }

    if (!text.contains("<i>") && !text.contains("</i>")) {
      return;
    }

    List<RunChunk> chunks = new ArrayList<>();
    StringBuilder chunkText = new StringBuilder();
    CTRPr chunkSourceRPr = null;
    boolean italic = false;
    boolean chunkItalic = false;

    int cursor = 0;
    while (cursor < text.length()) {
      if (text.startsWith("<i>", cursor)) {
        if (chunkText.length() > 0) {
          chunks.add(new RunChunk(chunkText.toString(), chunkSourceRPr, chunkItalic));
          chunkText.setLength(0);
        }
        italic = true;
        cursor += 3;
        continue;
      }

      if (text.startsWith("</i>", cursor)) {
        if (chunkText.length() > 0) {
          chunks.add(new RunChunk(chunkText.toString(), chunkSourceRPr, chunkItalic));
          chunkText.setLength(0);
        }
        italic = false;
        cursor += 4;
        continue;
      }

      XWPFRun sourceRun = runByCharIndex.get(cursor);
      if (chunkText.length() == 0) {
        chunkSourceRPr = copyRunProperties(sourceRun);
        chunkItalic = italic;
      } else if (!sameRunProperties(sourceRun, chunkSourceRPr) || italic != chunkItalic) {
        chunks.add(new RunChunk(chunkText.toString(), chunkSourceRPr, chunkItalic));
        chunkText.setLength(0);
        chunkSourceRPr = copyRunProperties(sourceRun);
        chunkItalic = italic;
      }

      chunkText.append(text.charAt(cursor));
      cursor++;
    }

    if (chunkText.length() > 0) {
      chunks.add(new RunChunk(chunkText.toString(), chunkSourceRPr, chunkItalic));
    }

    for (int i = para.getRuns().size() - 1; i >= 0; i--) {
      para.removeRun(i);
    }

    for (RunChunk chunk : chunks) {
      addTextRun(para, chunk.text, chunk.italic, chunk.sourceRPr);
    }
  }

  /** Add a text run with optional italic formatting. */
  private static void addTextRun(XWPFParagraph para, String text, boolean italic, CTRPr sourceRPr) {
    if (text == null || text.isEmpty()) {
      return;
    }

    XWPFRun run = para.createRun();
    if (sourceRPr != null) {
      run.getCTR().setRPr((CTRPr) sourceRPr.copy());
    }
    run.setItalic(italic);
    run.setText(text);
  }

  /** Copy run properties if present. */
  private static CTRPr copyRunProperties(XWPFRun run) {
    if (run == null || run.getCTR() == null || !run.getCTR().isSetRPr()) {
      return null;
    }
    return (CTRPr) run.getCTR().getRPr().copy();
  }

  /** Compare run properties object identity by underlying XML equality fallback. */
  private static boolean sameRunProperties(XWPFRun run, CTRPr rPr) {
    CTRPr runRPr = copyRunProperties(run);
    if (runRPr == null && rPr == null) {
      return true;
    }
    if (runRPr == null || rPr == null) {
      return false;
    }
    return runRPr.xmlText().equals(rPr.xmlText());
  }

  /** Piece of rebuilt paragraph text with source formatting and italic state. */
  private static class RunChunk {
    private final String text;
    private final CTRPr sourceRPr;
    private final boolean italic;

    private RunChunk(String text, CTRPr sourceRPr, boolean italic) {
      this.text = text;
      this.sourceRPr = sourceRPr;
      this.italic = italic;
    }
  }

  /** Bookmark spell
   * @param doc Word document
   * @param bookmarkName Name of the bookmark being added
   * @param text Text of the paragraph which is to be the start
   * and end of the marked text. */
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

  /** Index entry spell
   * @param para A paragraph in the document
   * @param entryText Index entry text. */
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

  /** Split heading spell
   * @param doc Word document
   * @param headingText Text to go into a level 2 heading
   * @param trailingText Text to follow the heading text in the paragraph.*/
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
   * The results can be odd if both strings are empty.
   * @param doc The document
   * @param headingText Text that will appear in TOC
   * If this is empty, the trailing text becomes the only paragraph added.
   * @param trailingText Additional text that won't appear in TOC
   */
  public static void addSplitHeading2Para(XWPFDocument doc,
      String headingText,
      String trailingText) {
    if ((headingText != null) && (headingText.length() > 0)) {
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
    }

    // Second paragraph with the trailing text (no heading style)
    XWPFParagraph trailingPara = doc.createParagraph();
    XWPFRun trailingRun = trailingPara.createRun();
    trailingRun.setText(" " + trailingText);
    applyItalicTags(trailingPara, doc);
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
   * Insert a borderless floating text box at the front of the paragraph, styled
   * like the paragraph but with doubled font size to mimic a drop-cap string.
   * The text box wraps with surrounding text.
   * @param para target paragraph
   * @param doc document containing styles
   * @param dropText text to show in the floating box
   */
  public static void dropTextBox(XWPFParagraph para, XWPFDocument doc, String dropText) {
    if (para == null || doc == null || dropText == null || dropText.isEmpty()) {
      return;
    }

    int baseHalfPoints = resolveParagraphStyleFontSizeHalfPoints(para, doc);
    int dropHalfPoints = Math.max(2, baseHalfPoints * 2);
    int dropPointSize = Math.max(1, dropHalfPoints / 2);

    double widthPt = Math.max(24.0, (dropText.length() * dropPointSize * 0.62) + 6.0);
    double heightPt = Math.max(18.0, (dropPointSize * 1.35));

    XWPFRun run = para.insertNewRun(0);
    CTPicture pict = run.getCTR().addNewPict();
    XmlCursor cursor = pict.newCursor();
    try {
      final String nsV = "urn:schemas-microsoft-com:vml";
      final String nsO = "urn:schemas-microsoft-com:office:office";
      final String nsW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
      final String nsW10 = "urn:schemas-microsoft-com:office:word";

      cursor.toEndToken();
      cursor.beginElement("shape", nsV);
      cursor.insertAttributeWithValue("id", "dropTextBox" + System.nanoTime());
      cursor.insertAttributeWithValue("type", "#_x0000_t202");
      cursor.insertAttributeWithValue(
          "style",
          "position:absolute;"
              + "margin-left:0pt;"
              + "margin-top:0pt;"
            + "width:" + String.format(java.util.Locale.ROOT, "%.2f", widthPt) + "pt;"
            + "height:" + String.format(java.util.Locale.ROOT, "%.2f", heightPt) + "pt;"
              + "z-index:251659264;"
              + "mso-position-horizontal:left;"
              + "mso-position-horizontal-relative:text;"
              + "mso-position-vertical:top;"
              + "mso-position-vertical-relative:line;"
            + "mso-wrap-style:square;");
      cursor.insertAttributeWithValue("stroked", "f");
      cursor.insertAttributeWithValue("filled", "f");

      cursor.beginElement("wrap", nsW10);
      cursor.insertAttributeWithValue("type", "square");
      cursor.toEndToken();

      cursor.beginElement("textbox", nsV);
      cursor.insertAttributeWithValue("style", "mso-fit-shape-to-text:t;inset:0,0,0,0");

      cursor.beginElement("txbxContent", nsW);
      cursor.beginElement("p", nsW);
      boolean hasDropTextStyle = hasStyle(doc, "DropText");
      String styleId = hasDropTextStyle ? "DropText" : para.getStyleID();
      if (styleId != null && !styleId.isEmpty()) {
        cursor.beginElement("pPr", nsW);
        cursor.beginElement("pStyle", nsW);
        cursor.insertAttributeWithValue("val", nsW, styleId);
        cursor.toEndToken();
        cursor.toEndToken();
      }

      cursor.beginElement("r", nsW);
      if (!hasDropTextStyle) {
        // Fallback: retain the existing doubled-size behavior when DropText style is absent.
        cursor.beginElement("rPr", nsW);
        cursor.beginElement("sz", nsW);
        cursor.insertAttributeWithValue("val", nsW, Integer.toString(dropHalfPoints));
        cursor.toEndToken();
        cursor.beginElement("szCs", nsW);
        cursor.insertAttributeWithValue("val", nsW, Integer.toString(dropHalfPoints));
        cursor.toEndToken();
        cursor.toEndToken();
      }

      cursor.beginElement("t", nsW);
      cursor.insertChars(dropText);
      cursor.toEndToken();

      cursor.toEndToken();
      cursor.toEndToken();
      cursor.toEndToken();
      cursor.toEndToken();
      cursor.toEndToken();
      cursor.toEndToken();

      cursor.beginElement("lock", nsO);
      cursor.insertAttributeWithValue("ext", nsV, "edit");
      cursor.insertAttributeWithValue("anchorlock", "t");
      cursor.toEndToken();

      cursor.toEndToken();
    } finally {
      cursor.dispose();
    }
  }

  /** Resolve paragraph style font size in half-points (hps), with fallback. */
  private static int resolveParagraphStyleFontSizeHalfPoints(XWPFParagraph para, XWPFDocument doc) {
    int fallback = 22; // 11pt

    String styleId = para.getStyleID();
    if (styleId == null || styleId.isEmpty()) {
      return fallback;
    }

    XWPFStyles styles = doc.getStyles();
    if (styles == null) {
      return fallback;
    }

    return resolveStyleFontSizeHalfPoints(styles, styleId, 0, fallback);
  }

  /** Check whether the document defines a style with the given id. */
  private static boolean hasStyle(XWPFDocument doc, String styleId) {
    if (doc == null || styleId == null || styleId.isEmpty()) {
      return false;
    }
    XWPFStyles styles = doc.getStyles();
    return styles != null && styles.getStyle(styleId) != null;
  }

  /** Resolve style font size recursively through basedOn chain. */
  private static int resolveStyleFontSizeHalfPoints(XWPFStyles styles,
      String styleId,
      int depth,
      int fallback) {
    if (styleId == null || styleId.isEmpty() || depth > 12) {
      return fallback;
    }

    XWPFStyle style = styles.getStyle(styleId);
    if (style == null || style.getCTStyle() == null) {
      return fallback;
    }

    CTStyle ctStyle = style.getCTStyle();
    if (ctStyle.isSetRPr()) {
      CTRPr rPr = ctStyle.getRPr();
      if (rPr.sizeOfSzArray() > 0 && rPr.getSzArray(0).getVal() != null) {
        Object val = rPr.getSzArray(0).getVal();
        try {
          return Integer.parseInt(val.toString());
        } catch (NumberFormatException e) {
          // fall through to basedOn/fallback
        }
      }
    }

    if (ctStyle.isSetBasedOn() && ctStyle.getBasedOn().getVal() != null) {
      return resolveStyleFontSizeHalfPoints(styles, ctStyle.getBasedOn().getVal(), depth + 1, fallback);
    }

    return fallback;
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

  /** Example save method
   * @param doc Word document
   * @param filename File name where it should be saved
   * @throws Exception when writing the file goes wrong
   */
  public static void saveDoc(XWPFDocument doc, String filename) throws Exception {
    try (FileOutputStream out = new FileOutputStream(filename)) {
      doc.write(out);
    }
  }

}
