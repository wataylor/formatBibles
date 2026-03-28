package asst.formatWord.snippets;

import java.io.BufferedReader;
import java.io.FileReader;

/**Check through all the Book files of two electronic bibles
 * to make sure that all book files always have the same chapter
 * and verse numbers
 * <\p><p>This program is rarely used so it does not produce an
 * executable .jar file.  Go to targets/classes and run
 * <code>java -cp . asst/formatWord/snippets/CompareElectronicBiblesMain</code>
 * and route the output someplace convenient.  Be sure to run
 * <code>maven clean install</code> on the pom so that any
 * changes go into target/classes.
 *
 * @author Material Gain
 * @since 2026 02
 */
public class CompareElectronicBiblesMain {

  /**
   * @param args //TODO make the arguments specify the 2 Bibles
   */
  public static void main(String[] args) {
    String path1 = "/Sync/Biblical/asciiBible/";
    String path2 = "/temp/KJB/kjv-bibleprotector-com/";
    int verseCount = 0;
    int wordDiffCount = 0;
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
          } else {
            if (compareWords(line1, line2)) {
              wordDiffCount++;
            }
          }
          lineCount++; verseCount++;
          line1 = reader1.readLine();
          line2 = reader2.readLine();
        } while ((line1 != null) && (line2 != null));
        reader1.close();
        reader2.close();
        /*System.out.println(f + " " + lineCount
            + " " + wordDiffCount + " differing words.");
         */
      }
    } catch (Exception e) {
      e.printStackTrace();
    }
    System.out.println("Checked " + verseCount + " verses in both folders."
	+ " " + wordDiffCount + " differing words.");
  }

  /** This program is passed 2 lines from 2 different electronic bibles.
   * They are known to have the same chapter and verse, so all the words
   * should be identical.  It is possible that they might include <i>
   * notations for Italics, so <i> and </i> must be stripped out before
   * comparing words.<p>
   * <p>The program goes through the 2 lines 1 word at a time.  If the words
   * do not have the same spelling, the program prints the book, chapter, and
   * verse followed by the two words.  If more than one word differs,
   * all words are printed on the same line.
   * @param line1
   * @param line2 This line may have notes enclosed in {@literal <<} {@literal >>}
   * If this is at the beginning, skip past {@literal >>} and the following space.
   * If it is not at the beginning skip until the space before {@literal <<}.
   */
  public static boolean compareWords(String line1, String line2) {
    int ix1 = line1.indexOf("  ");
    StringBuilder sb = new StringBuilder();
    /* This is known to match in both lines.*/
    String bookChapVerse = line1.substring(0, ix1);
    String words1 = line1.substring(ix1).replaceAll("<i>", "").replaceAll("</i>", "").trim();
    String words2 = line2.substring(ix1).replaceAll("<i>", "").replaceAll("</i>", "").trim();
    words1 = stripNotes(words1);
    words2 = stripNotes(words2);
    String[] w1 = words1.split("\\s+");
    String[] w2 = words2.split("\\s+");
    if (w1.length != w2.length) {
      System.out.println(bookChapVerse + " word count mismatch: \n" + words1 + "\n" + words2);
      return true;
    }
    for (int i = 0; i < w1.length; i++) {
      /* Record the differing words for this verse but not if
      they differ only in punctuation after each word. */
      String normalized1 = stripTrailingPunctuation(w1[i]);
      String normalized2 = stripTrailingPunctuation(w2[i]);
      if (!normalized1.equals(normalized2)) {
        sb.append(w1[i]).append("!=").append(w2[i]).append(" ");
      }
    }
    if (sb.length() > 0) {
      System.out.println(bookChapVerse + " " + sb.toString());
      return true;
    }
    return false;
  }

  /** A line may have notes enclosed in {@literal <<} {@literal >>}
   * If this is at the beginning, skip past {@literal >>} and the following space.
   * If it is not at the beginning skip until the space before {@literal <<}.
   * @param words a line of text with no italic formatting.
   * @return line with notes removed.
   */
  public static String stripNotes(String words) {
    int ix = words.indexOf("<<");
    if (ix < 0) { return words; }
    if (ix == 0) { /* Notes at the beginning << notes >>*/
      int iy = words.indexOf(">>");
      return words.substring(iy + 3); /* Skip >> and the space*/
    }
    return words.substring(0, ix-1);
  }

  /** Remove all punctuation from the end of a word
   * @param word
   * @return word or the word without ending punctuation.
   */
  public static String stripTrailingPunctuation(String word) {
    return word.replaceAll("\\p{Punct}+$", "");
  }

}
