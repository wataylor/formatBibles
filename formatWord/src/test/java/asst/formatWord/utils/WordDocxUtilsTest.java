package asst.formatWord.utils;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STVerticalAlignRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.junit.jupiter.api.*;

import java.math.BigInteger;

import static org.junit.jupiter.api.Assertions.*;

/** Test the methods in WordDocxUtils.  These methods change the underlying
 * .docx XML structure and check the structure to ensure that the expected
 * changes were made.
 * @author GitHub CoPilot
 * @since 2025 12
 */
public class WordDocxUtilsTest {

  private XWPFDocument doc;

  @BeforeEach
  public void setUp() {
    doc = new XWPFDocument();
  }

  @AfterEach
  public void tearDown() throws Exception {
    doc.close();
  }

  @Test
  public void testSuperscriptParagraph() {
    WordDocxUtils.addSuperscriptParagraph(doc, "1", " normal text");
    XWPFParagraph para = doc.getParagraphs().get(0);

    assertEquals("1 normal text", para.getText());
    assertEquals(STVerticalAlignRun.SUPERSCRIPT, para.getRuns().get(0).getVerticalAlignment());
  }

  @Test
  public void testFootnoteInserted() {
    XWPFParagraph para = doc.createParagraph();
    para.createRun().setText("Text with footnote");
    WordDocxUtils.addFootnote(para, doc, "Footnote text");

    assertFalse(doc.getFootnotes().isEmpty());
    XWPFFootnote footnote = doc.getFootnotes().get(0);
    assertEquals("Footnote text", footnote.getParagraphs().get(0).getText());
  }

  @Test
  public void testBookmarkInserted() {
    WordDocxUtils.addBookmarkParagraph(doc, "Chapter1Start", "Chapter 1");
    XWPFParagraph para = doc.getParagraphs().get(0);
    assertTrue(para.getCTP().getBookmarkStartList().stream()
	.anyMatch(b -> "Chapter1Start".equals(b.getName())));
  }

  // @Test
  public void testIndexEntryInserted() {
    XWPFParagraph para = doc.createParagraph();
    WordDocxUtils.addIndexEntry(para, "Sample Entry");
    boolean hasXE = para.getCTP().getRArray().length > 0 &&
        para.getCTP().getRArray(0).getFldCharList().size() > 0;
    assertTrue(hasXE, "Index entry field should be present");
     
  }

  @Test
  public void testSplitHeading() {
    WordDocxUtils.addSplitHeading(doc, "word1 word2", "word3 word4");

    assertEquals("Heading2", doc.getParagraphs().get(0).getStyle());
    assertEquals("Normal", doc.getParagraphs().get(1).getStyle());
    assertEquals("word1 word2", doc.getParagraphs().get(0).getText());
    assertEquals("word3 word4", doc.getParagraphs().get(1).getText());
  }

  @Test
  public void testHyperlinkToBookmark() {
    // Create a target bookmark
    WordDocxUtils.addBookmarkParagraph(doc, "target", "This is the target");

    // Create a paragraph with a hyperlink to that bookmark
    XWPFParagraph para = doc.createParagraph();
    WordDocxUtils.addHyperlinkToBookmark(para, "target", "source");

    // Verify the hyperlink anchor
    CTHyperlink hyperlink = para.getCTP().getHyperlinkArray(0);
    assertEquals("target", hyperlink.getAnchor());

    // Verify the visible text of the hyperlink
    String text = hyperlink.getRArray(0).getTArray(0).getStringValue();
    assertEquals("source", text);
  }

  @Test
  public void testMultipleHyperlinksInOneParagraph() {
    // Create multiple targets
    WordDocxUtils.addBookmarkParagraph(doc, "target1", "Target 1");
    WordDocxUtils.addBookmarkParagraph(doc, "target2", "Target 2");

    // Create paragraph with multiple hyperlinks
    XWPFParagraph para = doc.createParagraph();
    WordDocxUtils.addHyperlinkToBookmark(para, "target1", "source1");
    WordDocxUtils.addHyperlinkToBookmark(para, "target2", "source2");

    // Verify both hyperlinks exist
    assertEquals(2, para.getCTP().getHyperlinkList().size());
    assertEquals("target1", para.getCTP().getHyperlinkArray(0).getAnchor());
    assertEquals("target2", para.getCTP().getHyperlinkArray(1).getAnchor());
  }

}
