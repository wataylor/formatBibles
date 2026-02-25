package asst.formatWord.snippets;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.StyleDescription;
import org.apache.poi.hwpf.model.StyleSheet;

import java.io.FileInputStream;

/**  POI requires that styles be referenced either by their ID or by
 * name as Word stored it.  It proved impossible to get a list of styles
 * from a .docx file so this program was written to explore the .doc
 * version of the document.  The way POI specifies styles in .doc and
 * in .docx is not the same.
 * 
 * @author Material Gain
 * @since 2015 12
 */
public class ListStylesDocMain {

  /**
   * @param args One path to a .doc file
   * @throws Exception
   */
  public static void main(String[] args) throws Exception {
    try (FileInputStream fis = new FileInputStream(args[0])) {
      HWPFDocument doc = new HWPFDocument(fis);
      StyleSheet styleSheet = doc.getStyleSheet();

      int numStyles = styleSheet.numStyles();
      System.out.println("Total styles: " + numStyles);

      for (int i = 0; i < numStyles; i++) {
	StyleDescription style = styleSheet.getStyleDescription(i);
	if (style != null) {
	  System.out.println("Index: " + i + " | Name: " + style.getName());
	}
      }
      doc.close();
    } catch (Exception e) {
      e.printStackTrace();
    }
  }
}
