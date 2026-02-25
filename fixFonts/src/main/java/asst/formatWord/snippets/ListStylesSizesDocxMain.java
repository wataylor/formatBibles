package asst.formatWord.snippets;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyles;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;

import java.io.FileInputStream;
import java.lang.reflect.Field;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** It turned out not to be possible to get a list of styles without
 * using reflection to make some private methods available.  The Style ID is
 * the name by which POI assigns the style to paragraphs; the Name is
 * the display name.</p>
 * <p>Knowing what a style is based on minimizes the number of sizes
 * that must be set.
 * 
 * @author Material Gain
 * @since 2015 12
 */
public class ListStylesSizesDocxMain {
  /**
   * @param args One full path to a .docx file
   * @throws Exception
   */
  public static void main(String[] args) throws Exception {
    if (args.length == 0) {
      System.out.println("Usage: ListStylesSizesDocxMain <path-to-docx>");
      return;
    }
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
	  System.out.println("Type\tID\tName\tBasedOn\tSize");

	  for (CTStyle style : styleList) {
	    String styleId = style.getStyleId();
	    String name = (style.getName() != null) ? style.getName().getVal() : "(no name)";
	    String type = (style.getType() != null) ? style.getType().toString() : "(no type)";
	    String basedOn = (style.getBasedOn() != null) ? style.getBasedOn().getVal() : "(none)";
	    String fontSize = extractFontSize(style);
	    System.out.println(type + "\t" + styleId + "\t" + name + "\t" + basedOn + "\t" + fontSize);
	  }
	}
      }
      doc.close();
    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private static String extractFontSize(CTStyle style) {
    // Word stores size in half-points (e.g., 24 = 12pt)
    String xml = style.xmlText();

    // Prefer w:sz (latin) then w:szCs (complex script)
    String size = extractHalfPoints(xml, "w:sz");
    if (size == null) {
      size = extractHalfPoints(xml, "w:szCs");
    }

    if (size == null) {
      return "(no size)";
    }

    int halfPoints = Integer.parseInt(size);
    if (halfPoints > 0 && halfPoints % 2 == 0) {
      return (halfPoints / 2) + "pt";
    }
    if (halfPoints > 0) {
      return (halfPoints / 2.0) + "pt";
    }
    return "(no size)";
  }

  private static String extractHalfPoints(String xml, String elementName) {
    Pattern szPattern = Pattern.compile("<" + elementName + "[^>]*w:val=\\\"(\\d+)\\\"");
    Matcher matcher = szPattern.matcher(xml);
    if (matcher.find()) {
      return matcher.group(1);
    }
    return null;
  }
}
