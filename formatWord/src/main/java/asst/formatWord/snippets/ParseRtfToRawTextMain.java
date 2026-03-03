package asst.formatWord.snippets;

import java.io.File;
import java.io.FileInputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

import javax.swing.text.AttributeSet;
import javax.swing.text.Element;
import javax.swing.text.StyleConstants;
import javax.swing.text.StyledDocument;
import javax.swing.text.rtf.RTFEditorKit;

import asst.common.DescribeArgs;
import asst.common.MainArgs;

/** This program reads a RTF file and writes text.
 * @author Material Gain
 * @since 2026 02
 */
public class ParseRtfToRawTextMain {

  /** Describe the purpose of the command line args for the Help function.  */
  public static Map<String, String> argDescs = new HashMap<String, String>();

  static {
    argDescs.put("help", "If \"+help\" is specified, nothing else is run."
        + " Enter -help to turn help off to run the program.");
    argDescs.put("doIt", "\"+doIt\" must be set to take any action."
	+ "  If it is not set, no action is taken and information about"
	+ " actions that would be taken is printed instead.");
    argDescs.put("verbose", "\"+verbose\" increases the printout volume.");
    argDescs.put("inputFile", "RTF file to be converted.");
    argDescs.put("outputFile", "File where the converted file is written. ");
  }

  /** +help is the default value so that the program explains
   * the parameters if it is called with no arguments. */
  static final String[] DEFAULT_ARGS = {
      "-doIt", "-verbose",
      "inputFile=/Sync/Biblical/KJV/KJB-PCE-RTF.rtf",
      "outputFile=/temp/KJB/KJB-PCE-RTF.txt",
      "+help",
  };

  /** Convert a .rtf file to a string
   * @param rtfFile input file
   * @return output string
   * @throws Exception when things to wrong
   */
  public static String convert(File rtfFile) throws Exception {
    RTFEditorKit kit = new RTFEditorKit();
    StyledDocument doc = (StyledDocument) kit.createDefaultDocument();
    kit.read(new FileInputStream(rtfFile), doc, 0);

    StringBuilder out = new StringBuilder();

    boolean inItalic = false;
    boolean inBold = false;
    boolean inUnderline = false;

    for (int i = 0; i < doc.getLength(); i++) {
      Element elem = doc.getCharacterElement(i);
      AttributeSet as = elem.getAttributes();

      boolean italic = StyleConstants.isItalic(as);
      boolean bold = StyleConstants.isBold(as);
      boolean underline = StyleConstants.isUnderline(as);

      // Italic transitions
      if (italic && !inItalic) { out.append("<i>"); inItalic = true; }
      if (!italic && inItalic) { out.append("</i>"); inItalic = false; }

      // Bold transitions
      if (bold && !inBold) { out.append("<b>"); inBold = true; }
      if (!bold && inBold) { out.append("</b>"); inBold = false; }

      // Underline transitions
      if (underline && !inUnderline) { out.append("<u>"); inUnderline = true; }
      if (!underline && inUnderline) { out.append("</u>"); inUnderline = false; }

      // Append the character
      out.append(doc.getText(i, 1));
    }

    // Close any open tags
    if (inItalic) out.append("</i>");
    if (inBold) out.append("</b>");
    if (inUnderline) out.append("</u>");

    return out.toString();
  }

  /**
   * @param args as described in the argDescs map
   */
  public static void main(String[] args) {
    MainArgs carg = new MainArgs(DEFAULT_ARGS);
    carg.parseArgs(args); // updates or overrides the default values

    if (carg.getBoolean("help")) {
      StringBuilder sb =
	  DescribeArgs.describeArgs(carg, DEFAULT_ARGS, argDescs,
	      "There may be NO spaces before or after the = in a string parameter value.\n"
		  + "If a value has spaces, the entire string must be enclosed in quotes or the space must be escaped.\n");
      System.out.print(sb.toString());
      System.exit(0);
    }

    String inputFile = (String)carg.get("inputFile");
    String outputFile = (String)carg.get("outputFile");

    try {
      File inFile = new File(inputFile);
      Path outPath = Paths.get(outputFile);
      String result= convert(inFile);
      Files.write(outPath, result.getBytes("UTF-8"));
    } catch (Exception e) {
      e.printStackTrace();
    }

  }

}
