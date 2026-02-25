package asst.formatWord.snippets;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;

import java.io.FileInputStream;
import java.util.List;

/** List all the bookmarks in a document and the paragraphs
 * where they are defined.
 * This helps generate URLs which use the bookmarks as anchors.
 * Excludes bookmarks whose names begin with _Toc because they
 * are automatically generated and change with the TOC.
 * @author GitHub GoPilot
 * @since 2025 12
 */
public class ListBookmarksDocxMain {
  /**
   * @param args One name of a .docx file
   * @throws Exception
   */
  public static void main(String[] args) throws Exception {
    if (args.length < 1) {
      System.err.println("Usage: ListBookmarksDocxMain <docx-file>");
      System.exit(1);
    }

    XWPFDocument doc = null;
    try (FileInputStream fis = new FileInputStream(args[0])) {
      doc = new XWPFDocument(fis);

      List<XWPFParagraph> paragraphs = doc.getParagraphs();

      for (XWPFParagraph para : paragraphs) {
	List<CTBookmark> bookmarks = para.getCTP().getBookmarkStartList();

	for (CTBookmark bookmark : bookmarks) {
	  String bookmarkName = bookmark.getName();
	  String paragraphText = para.getText();
	  if (bookmarkName.startsWith("_Toc")) { continue; }
	  // Truncate long paragraphs for readability
	  if (paragraphText.length() > 100) {
	    paragraphText = paragraphText.substring(0, 100) + "...";
	  }
	  System.out.println(bookmarkName + "\t" + paragraphText);
	}
      }
    } catch (Exception e) {
      e.printStackTrace();
    } finally {
      try {
	doc.close();
      } catch (Exception e) {
	System.out.println("Error closing doc file " + e.getMessage());
      }
    }
  }
}
