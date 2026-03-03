package asst.formatWord.snippets;

import java.io.BufferedReader;
import java.io.FileReader;

/**Check through all the Book files of two electronic bibles
 * to make sure that all book files always have the same chapter
 * and verse numbers
 * @author Material Gain
 * @since 2026 02
 */
public class CompareElectronicBiblesMain {

  public static void main(String[] args) {
    String path1 = "/Sync/Biblical/asciiBible/";
    String path2 = "/temp/KJB/kjv-bibleprotector-com/";
    int verseCount = 0;
    try {
      for (String f : TxtKJBTo66ChapterFilesMain.BOOK_FILE_NAMES) {
        if (f.length() < 4) { continue; }
        BufferedReader reader1 = new BufferedReader(new FileReader(path1 + f));
        BufferedReader reader2 = new BufferedReader(new FileReader(path2 + f));
        String line1 = reader1.readLine();
        String line2 = reader2.readLine();
        int lineCount = 0;
        int bad = 0;
        do {
          int ix1 = line1.indexOf("  ");
          int ix2 = line2.indexOf("  ");
          if (!line1.substring(0, ix1).equals(line2.substring(0, ix2))) {
            System.out.println("ERR " + f + " "
        	+ line1.substring(0, ix1)
        	+ " " + line2.substring(0, ix2));
            if (++bad > 5) {
              break;
            }
          }
          lineCount++; verseCount++;
          line1 = reader1.readLine();
          line2 = reader2.readLine();
        } while ((line1 != null) && (line2 != null));
        System.out.println(f + " " + lineCount);
      }
    } catch (Exception e) {
       e.printStackTrace();
    }
    System.out.println("Checked " + verseCount + " verses in both folders.");
  }

}
