package asst.formatWord.snippets;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;

/** Print bookmarks in a .docx file and all the paragraphs which reference them.
 * This helps generate URLs which use bookmarks as anchors.
 * Excludes bookmarks whose names begin with _Toc because they
 * are automatically generated and change with the TOC.
 * @author GitHub GoPilot
 * @since 2025 12
 */
public class ListBookmarkReferencesMain {
  /**
   * @param args One name of a .docx file
   * @throws Exception
   */
  public static void main(String[] args) throws Exception {
    if (args.length < 1) {
      System.err.println("Usage: ListBookmarkReferencesMain <docx-file>");
      System.exit(1);
    }

    XWPFDocument doc = null;
    try (FileInputStream fis = new FileInputStream(args[0])) {
      doc = new XWPFDocument(fis);

      // First, collect all bookmark names (excluding _Toc bookmarks)
      Set<String> bookmarkNames = new TreeSet<>();
      for (XWPFParagraph para : doc.getParagraphs()) {
	for (CTBookmark bookmark : para.getCTP().getBookmarkStartList()) {
	  String name = bookmark.getName();
	  if (name != null && !name.startsWith("_Toc")) {
	    bookmarkNames.add(name);
	  }
	}
      }

      // For each bookmark, find all references to it
      for (String bookmarkName : bookmarkNames) {
	List<String> references = new ArrayList<>();

	for (XWPFParagraph para : doc.getParagraphs()) {
	  List<CTHyperlink> hyperlinks = para.getCTP().getHyperlinkList();

	  for (CTHyperlink hyperlink : hyperlinks) {
	    String anchor = hyperlink.getAnchor();
	    if (bookmarkName.equals(anchor)) {
	      String paragraphText = para.getText();

	      // Truncate long paragraphs for readability
	      if (paragraphText.length() > 100) {
		paragraphText = paragraphText.substring(0, 100) + "...";
	      }

	      references.add(paragraphText);
	    }
	  }
	}

	// Print bookmark and its references
	if (!references.isEmpty()) {
	  for (String refText : references) {
	    System.out.println(bookmarkName + "\t" + refText);
	  }
	} else {
	  System.out.println(bookmarkName + "\tHas no references.");
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
