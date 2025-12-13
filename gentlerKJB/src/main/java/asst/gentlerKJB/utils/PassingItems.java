package asst.gentlerKJB.utils;

/** Class to hold a line of text and words being manipulated within it.
 * One instance is used for all the words in the list.
 * @author Material Gain
 * @since 2025 12
 */
public class PassingItems {
  /** A 2-digit book number to sort the verse references*/
  public String bkno;
  /** The incoming text line in lower case*/
  public String lineLowerCase;
  /** The original incoming text line */
  public String line;
  /** The word to be matched in lower case */
  public String oldWord;
  /** The word which replaces the old word in lower case.  It will
   * get an initial cap if the word was capitalized in the incoming
   * string. */
  public String newWord;
  /** Remember the chapter and verse for the line.  */
  public String bookChapVerse;
  /** Set true if anything changed */
  public boolean isDirty = false;

  /** Information about an input line, an old word, and the replacement
   * new word.
   * @param line original line of text
   * @param oldWord word which will be replaced
   * @param newWord word which will replace the old word 
   */
  public PassingItems(String line, String oldWord, String newWord) {
    /* Line begins with 3 letter book name, space, chapter : verse and 
     * 2 spaces */
    int ix = line.indexOf("  ");
    if (ix > 4) {
      this.bookChapVerse = line.substring(0, ix);
      this.line = line.substring(ix+2);
    } else {
      /* This makes testing a bit easier.  */
      this.bookChapVerse = "None. ";
      this.line = line;
    }
    this.lineLowerCase = this.line.toLowerCase();
    this.oldWord = oldWord.toLowerCase().trim();
    this.newWord = newWord.toLowerCase().trim();
  }
 
  /** Set the words for the next pass through the verse
   * @param oldWord archaic word
   * @param newWord replacement word
   */
  public void setWords(String oldWord, String newWord) {
    this.oldWord = oldWord.toLowerCase().trim();
    this.newWord = newWord.toLowerCase().trim();
  }

  /**
   * @return the edited line with any required alterations
   * made. 
   */
  public String getEditedLine() {
    return bookChapVerse + "  " + line.replaceAll("   ", " ");
  }

}
