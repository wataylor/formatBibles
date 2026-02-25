package asst.formatWord.snippets;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.namespace.QName;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyles;

/** This program reads a .docx file and sets the font sizes of listed
 * styles from an Excel .csv file.
 * @author Material gain
 * @since 2026 02
 */
public class SetStyleFontSizesMain {

  private static final String W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

  private static class StyleSize {
    final String sizeText;
    final boolean noSize;

    StyleSize(String sizeText, boolean noSize) {
      this.sizeText = sizeText;
      this.noSize = noSize;
    }
  }

  /** The .csv file has headers.  The important headers are:
   * ID The name by which Word identifies the style when editing
   * its properties.  Size12, Size10, Size8 tells which column
   * to use to set style sizes in the .docx.</p>
   * <p>The program walks the styles, setting the size based on the
   * matching style ID.  If the size is a number, set it to that many
   * pixels.  If it is (no size), set the size to have no value.</p>
   * <p>It is an error for a paragraph style with no size not to
   * inherit from another style which does have a size. 
   * @param args .docx file, .csv file, and which column has the
   * desired sizes.
   * @throws Exception when things go wrong
   */
  public static void main(String[] args) throws Exception {
    if (args.length < 3) {
      System.out.println("Usage: SetStyleFontSizesMain <docx> <csv> <SizeColumn>");
      return;
    }

    Path docxPath = Paths.get(args[0]);
    Path csvPath = Paths.get(args[1]);
    String sizeColumn = args[2];

    Map<String, StyleSize> styleSizes = readSizesFromCsv(csvPath, sizeColumn);
    if (styleSizes.isEmpty()) {
      System.out.println("No styles found for column: " + sizeColumn);
      return;
    }

    try (FileInputStream fis = new FileInputStream(docxPath.toFile());
         XWPFDocument doc = new XWPFDocument(fis)) {

      XWPFStyles styles = doc.getStyles();
      if (styles == null) {
        System.out.println("No styles found in document.");
        return;
      }

      java.lang.reflect.Field ctStylesField = XWPFStyles.class.getDeclaredField("ctStyles");
      ctStylesField.setAccessible(true);
      CTStyles ctStyles = (CTStyles) ctStylesField.get(styles);
      if (ctStyles == null) {
        System.out.println("No CTStyles found in document.");
        return;
      }

      List<CTStyle> styleList = ctStyles.getStyleList();
      int updated = 0;

      for (CTStyle style : styleList) {
        String styleId = style.getStyleId();
        if (styleId == null) {
          continue;
        }

        StyleSize size = styleSizes.get(styleId);
        if (size == null) {
          continue;
        }

        if (size.noSize) {
          removeFontSize(style);
        } else {
          Integer halfPoints = toHalfPoints(size.sizeText);
          if (halfPoints != null) {
            setFontSize(style, halfPoints);
          }
        }

        updated++;
      }

      try (FileOutputStream fos = new FileOutputStream(docxPath.toFile())) {
        doc.write(fos);
      }

      System.out.println("Updated styles: " + updated + ". Saved: " + docxPath);
    }
  }

  private static Map<String, StyleSize> readSizesFromCsv(Path csvPath, String sizeColumn) throws Exception {
    List<String> lines = Files.readAllLines(csvPath, StandardCharsets.UTF_8);
    Map<String, StyleSize> sizes = new HashMap<String, StyleSize>();

    if (lines.isEmpty()) {
      return sizes;
    }

    List<String> headers = parseCsvLine(lines.get(0));
    int idIndex = indexOfHeader(headers, "ID");
    int sizeIndex = indexOfHeader(headers, sizeColumn);

    if (idIndex < 0 || sizeIndex < 0) {
      System.out.println("Missing required headers. Found: " + headers);
      return sizes;
    }

    for (int i = 1; i < lines.size(); i++) {
      if (lines.get(i).trim().isEmpty()) {
        continue;
      }

      List<String> row = parseCsvLine(lines.get(i));
      if (row.size() <= Math.max(idIndex, sizeIndex)) {
        continue;
      }

      String id = row.get(idIndex).trim();
      if (id.isEmpty()) {
        continue;
      }

      String sizeValue = row.get(sizeIndex).trim();
      if (sizeValue.isEmpty()) {
        continue;
      }

      boolean noSize = sizeValue.equalsIgnoreCase("(no size)");
      sizes.put(id, new StyleSize(sizeValue, noSize));
    }

    return sizes;
  }

  private static int indexOfHeader(List<String> headers, String header) {
    for (int i = 0; i < headers.size(); i++) {
      if (headers.get(i).trim().equalsIgnoreCase(header)) {
        return i;
      }
    }
    return -1;
  }

  private static List<String> parseCsvLine(String line) {
    List<String> result = new ArrayList<String>();
    StringBuilder current = new StringBuilder();
    boolean inQuotes = false;

    for (int i = 0; i < line.length(); i++) {
      char c = line.charAt(i);
      if (c == '"') {
        if (inQuotes && i + 1 < line.length() && line.charAt(i + 1) == '"') {
          current.append('"');
          i++;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (c == ',' && !inQuotes) {
        result.add(current.toString());
        current.setLength(0);
      } else {
        current.append(c);
      }
    }

    result.add(current.toString());
    return result;
  }

  private static Integer toHalfPoints(String value) {
    String cleaned = value.trim().toLowerCase().replace("pt", "");
    try {
      double points = Double.parseDouble(cleaned);
      if (points <= 0) {
        return null;
      }
      return (int) Math.round(points * 2.0);
    } catch (NumberFormatException e) {
      return null;
    }
  }

  private static void setFontSize(CTStyle style, int halfPoints) {
    if (style.getRPr() == null) {
      style.addNewRPr();
    }

    setOrCreateSizeElement(style.getRPr(), "sz", halfPoints);
    setOrCreateSizeElement(style.getRPr(), "szCs", halfPoints);
  }

  private static void removeFontSize(CTStyle style) {
    if (style.getRPr() == null) {
      return;
    }

    removeSizeElements(style.getRPr(), "sz");
    removeSizeElements(style.getRPr(), "szCs");
  }

  private static void setOrCreateSizeElement(XmlObject rPr, String elementName, int halfPoints) {
    XmlObject[] existing = rPr.selectPath(
        "declare namespace w='" + W_NS + "' .//w:" + elementName);

    if (existing != null && existing.length > 0) {
      for (XmlObject obj : existing) {
        setSizeAttribute(obj, halfPoints);
      }
      return;
    }

    XmlCursor cursor = rPr.newCursor();
    cursor.toEndToken();
    cursor.beginElement(new QName(W_NS, elementName));
    cursor.insertAttributeWithValue(new QName(W_NS, "val"), String.valueOf(halfPoints));
    cursor.dispose();
  }

  private static void removeSizeElements(XmlObject rPr, String elementName) {
    XmlObject[] existing = rPr.selectPath(
        "declare namespace w='" + W_NS + "' .//w:" + elementName);
    if (existing == null) {
      return;
    }

    for (XmlObject obj : existing) {
      XmlCursor cursor = obj.newCursor();
      cursor.removeXml();
      cursor.dispose();
    }
  }

  private static void setSizeAttribute(XmlObject szObj, int halfPoints) {
    XmlCursor cursor = szObj.newCursor();
    cursor.setAttributeText(new QName(W_NS, "val"), String.valueOf(halfPoints));
    cursor.dispose();
  }

}
