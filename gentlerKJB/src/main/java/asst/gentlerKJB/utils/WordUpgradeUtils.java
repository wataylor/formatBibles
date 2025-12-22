package asst.gentlerKJB.utils;

import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** Utilities to help replace archaic words with more modern words
 * @author Material Gain
 * @since 2025 12
 */
public class WordUpgradeUtils {

  /** Map which holds a string builder for each archaic word
   * which accumulates chapter and verse references where the
   * archaic word was changed.*/
  public static Map<String, StringBuilder> wordChanges =
      new TreeMap<String, StringBuilder>();
  /** Set which holds all chapter and verse references which were
   * changed.  It is a tree set to make it easier to dump in 
   * canonical order.  */
  public static Set<String> verseChanges = new TreeSet<String>();

  /** return the index of where a word is found in a line unless
   * the word is found somewhere within a air of [].  If a ]
   * is found starting at the point where the word was found, 
   * it is enclosed in [] so does not count as being found. 
   * @param line a line of text
   * @param word the word or phrase being sought
   * @return indexOf the word in the line or -1 if it is in []
   */
  public static int findWordIndex(String line, String word) {
    int ix = line.indexOf(word);
    if (ix < 0) { // did not find it, this is by far the usual case
      return ix;
    }
    /* The word was found somewhere, but it might not be a
     * stand alone word.  Given that it is in the line, it is
     * worth the time to create a pattern */
    Pattern pattern = Pattern.compile("\\b" + Pattern.quote(word) + "\\b");
    Matcher matcher = pattern.matcher(line);
    if (matcher.find()) {
        ix = matcher.start();  // Position of match
    } else {
      return -1;
    }
    /* The word is not found if it is found anywhere inside [] */
    int iy = line.indexOf("]", ix);
    if (iy < 0) { return ix;}
    int iz = line.indexOf("[", ix);
    if (iz < 0) { return -1; } // no [ after the word
    if (iz < iy) { return ix; }
    return -1;
  }

  /** Edit a line to replace an archaic word with a new word and
   * include the archaic word in [].  Record chapter and verse for
   * both if it changed.
   * @param pi 
   * @return true if the line changed because of this word.
   */
  public static void modernizeWord(PassingItems pi) {
    int ix = findWordIndex(pi.lineLowerCase, pi.oldWord);
    if (ix >= 0) {
      String text = pi.newWord;
      if (Character.isUpperCase(pi.line.charAt(ix))) {
	/* Capitalize the word as the input word was capitalized*/
	text = text.substring(0, 1).toUpperCase() + text.substring(1);
      }
      pi.line = pi.line.substring(0, ix) + text + " ["
	  + pi.line.substring(ix, ix + pi.oldWord.length()) + "]"
	  + pi.line.substring(ix + pi.oldWord.length());
      pi.lineLowerCase = pi.line.toLowerCase();
      recordCref(pi);
    }
  }

  /** replace oldWord with newWord but do not retain oldWord in []
   * This is used to modernize spelling as in honour becomes honor.
   * @param pi
   */
  public static void replaceWord(PassingItems pi) {
    int ix = -1;
    do {
      ix = findWordIndex(pi.lineLowerCase, pi.oldWord);
      if (ix < 0) { return; }
      String text = pi.newWord;
      if (Character.isUpperCase(pi.line.charAt(ix))) {
	/* Capitalize the word as the input word was capitalized*/
	text = text.substring(0, 1).toUpperCase() + text.substring(1);
      }
      pi.line = pi.line.substring(0, ix) + text
	  + pi.line.substring(ix + pi.oldWord.length());
      ix += pi.oldWord.length();
      pi.lineLowerCase = pi.line.toLowerCase();
      recordCref(pi);
    } while (ix < pi.line.length()-1);
  }

  /** Note the change to the verse in the global list and the
   * change on behalf of the old word.
   * @param pi
   */
  public static void recordCref(PassingItems pi) {
    String change = pi.bkno + pi.bookChapVerse;
    verseChanges.add(change);
    String wordKey = pi.oldWord + " -> " + pi.newWord;
    StringBuilder sb = wordChanges.get(wordKey);
    if (sb == null) {
	sb = new StringBuilder();
	wordChanges.put(wordKey, sb);
    }
    if (sb.indexOf(change) < 0) { // might replace more than once
      sb.append(change + "_");
    }
    pi.isDirty = true;
  }
}
