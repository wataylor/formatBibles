package asst.formatWord.snippets;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyles;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;

import java.io.FileInputStream;
import java.lang.reflect.Field;
import java.util.List;

/** It turned out not to be possible to get a list of styles without
 * using reflection to make some private methods available.  The Style ID is
 * the name by which POI assigns the style to paragraphs; the Name is
 * the display name.
 * 
 * @author Material Gain
 * @since 2015 12
 */
public class ListStylesDocxReflectionMain {
  /**
   * @param args One full path to a .docx file
   * @throws Exception
   */
  public static void main(String[] args) throws Exception {
    try (FileInputStream fis = new FileInputStream(args[0])) {
      XWPFDocument doc = new XWPFDocument(fis);
      XWPFStyles styles = doc.getStyles();

      if (styles != null) {
	// Use reflection to access the private ctStyles field
	Field ctStylesField = XWPFStyles.class.getDeclaredField("ctStyles");
	ctStylesField.setAccessible(true);
	CTStyles ctStyles = (CTStyles) ctStylesField.get(styles);

	if (ctStyles != null) {
	  List<CTStyle> styleList = ctStyles.getStyleList();
	  System.out.println("Total styles: " + styleList.size());

	  for (CTStyle style : styleList) {
	    String styleId = style.getStyleId();
	    String name = (style.getName() != null) ? style.getName().getVal() : "(no name)";
	    String type = (style.getType() != null) ? style.getType().toString() : "(no type)";
	    System.out.println("Type: " + type + " | StyleId: " + styleId + " | Name: " + name);
	  }
	}
      }
      doc.close();
    } catch (Exception e) {
      e.printStackTrace();
    }
  }
}
