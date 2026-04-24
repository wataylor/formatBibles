package asst.formatWord.snippets;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;

/**Go through two electronic bibles, one of which has paragraph marks
 * and the other does not.  Insert paragraph marks in the Bible
 * that does not have them.
 * <\p><p>This program is rarely used so it does not produce an
 * executable .jar file.  Go to targets/classes and run
 * <code>java -cp . asst/formatWord/snippets/AddRTFParamarksToTextKJVMain</code>
 * and route the output someplace convenient.  Be sure to run
 * <code>maven clean install</code> on the pom so that any
 * changes go into target/classes.
 *
 * @author Material Gain
 * @since 2026 04
 */

public class AddRTFParamarksToTextKJVMain {

  /** Generates more print if true.  */
  public static boolean verbose = false;
  /**
   * @param args //TODO make the arguments specify the 2 Bibles
   * and the output
   */
  public static void main(String[] args) {
    /* This file has the paragraph marks. It was made from the .rtf
     * which has them.*/
    String path1 = "/Sync/Biblical/KJV/KJB-PCE-RTF.txt";
    /* This file has the same verses but not paragraph marks*/
    String path2 = "/Sync/Biblical/KJV/KJB-PCE-127.txt";
    /* This file has the same verses but not paragraph marks*/
    String path3 = "/Temp/KJB/KJB-PCE-127-With.txt";
    /* These variables are defined outside the try loop so
     * they can be printed in the catch statement should that become
     * necessary.*/
    int paraCount = 0;
    String paraLine = "";
    int lineCount = 0;
    int inputLineCount = 0;
    String targetText;
    String noParaLine = "";
    String outLine = "";
    boolean done = false;
    int ix = 0;
    int iy = 0;

    try {
      BufferedReader reader1 = new BufferedReader(new FileReader(path1));
      BufferedReader reader2 = new BufferedReader(new FileReader(path2));
      BufferedWriter writer1 =
	  new BufferedWriter(
	      new OutputStreamWriter(
		  new FileOutputStream(path3),
		  StandardCharsets.UTF_8)
	      );
      paraLine = reader1.readLine();
      inputLineCount++;
      if (verbose) { System.out.println(inputLineCount + " " + paraLine); }
      do {
	ix = paraLine.indexOf("¶");
	if (ix > 0) {
	  targetText = paraLine.substring(ix+2)
	      .replace("<i>", "[")
	      .replace("</i>", "]")
	      .replace("] [", " ")
	      .replace("’", "'");
	  done = false;
	  do {
	    noParaLine = reader2.readLine();
	    if (noParaLine == null) {
	      System.out.println("Shriek! inputLine " + inputLineCount
		  + " ix " + ix
		  + " seeking\n" + targetText 
		  + "<\n" + paraLine + "<");
	      bail(reader1, reader2, writer1);
	    }
	    lineCount++;
	    if ( (iy = endsWithNoCase(noParaLine, targetText)) > 0) {
	      outLine = noParaLine.substring(0, iy);
	      outLine += "¶ " +  targetText;
	      writer1.append(outLine);
	      writer1.newLine();
	      paraCount++;
	      done = true;
	      targetText = null;
	    } else {
	      writer1.append(noParaLine);
	      writer1.newLine();
	    }
	  } while (!done);
	}
	paraLine = reader1.readLine();
	inputLineCount++;
	if (verbose) {
	  System.out.println(inputLineCount + " " + paraLine);
	  if (inputLineCount > 100) {
	    bail(reader1, reader2, writer1);
	  }
	}
      } while (paraLine != null);
      /* Copy the rest of reader 2 to writer 1.*/
      noParaLine = reader2.readLine();
      do {
	lineCount++;
	writer1.append(noParaLine);
	writer1.newLine();
	noParaLine = reader2.readLine();
      } while (noParaLine != null);
      System.out.println("Wrote " + lineCount + " lines, added "
	  + paraCount + " paragraphs.");
      bail(reader1, reader2, writer1);
    } catch (Exception e) {
      System.out.println(e.getMessage());
      e.printStackTrace();
    }
  }

  /** do an Ends With that is case-independent and return he index.
   * @param input main string
   * @param with what it might end with
   * @return Index of the with string in the input string.
   */
  public static int endsWithNoCase(String input, String with) {
    String inLow = input.toLowerCase().replace("-", "").replace("]", "")
	.replace("[", "");
    String withLow = with.toLowerCase().replace("-", "").replace("]", "")
	.replace("[", "");
    return inLow.indexOf(withLow);
  }

  /**
   * @param reader1
   * @param reader2
   * @param writer1
   * @throws IOException
   */
  public static void bail(BufferedReader reader1, BufferedReader reader2, BufferedWriter writer1) throws IOException {
    reader1.close();
    reader2.close();
    writer1.close();
    System.exit(0);
  }

}
