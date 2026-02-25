package asst.formatWord.snippets;

import java.io.FileInputStream;
import java.lang.reflect.Method;
import java.util.Arrays;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFStyles;

/** Uses reflection to find out the POI style methods
 * 
 * @author Material Gain
 * @since 2015 12
 */
public class DiscoverStylesMethodsMain {
  /**
   * @param args A path to a .docx file
   * @throws Exception
   */
  public static void main(String[] args) throws Exception {
    try (FileInputStream fis = new FileInputStream(args[0])) {
      XWPFDocument doc = new XWPFDocument(fis);
      // XWPFStyles styles = doc.getStyles();

      System.out.println("Available methods on XWPFStyles:");
      for (Method method : XWPFStyles.class.getMethods()) {
	if (method.getName().toLowerCase().contains("style") || 
	    method.getName().toLowerCase().contains("ct")) {
	  System.out.println("  " + method.getName() + 
	      " - params: " + Arrays.toString(method.getParameterTypes()) +
	      " - returns: " + method.getReturnType().getSimpleName());
	}
      }
      doc.close();
    } catch (Exception e) {
      e.printStackTrace();
    }
  }
}
