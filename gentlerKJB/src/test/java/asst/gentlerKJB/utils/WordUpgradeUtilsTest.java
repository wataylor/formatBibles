package asst.gentlerKJB.utils;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

class WordUpgradeUtilsTest {
  public static final String aLine = "word1 word2 word5a [word5] word3";
  public static final String g11 = "GEN 1:1  In the beginning God created the heaven and the earth.";
  public static final String mat12 = "MAT 1:2  Abraham Begat Isaac; and Isaac begat Jacob; and Jacob begat Judas and his brethren;";
  public static final String mat12New = "MAT 1:2  Abraham Fathered [Begat] Isaac; and Isaac begat Jacob; and Jacob begat Judas and his brethren;";
  public static final String jon1621 = "JOH 16:21  A woman when she is in travail hath sorrow, because her hour is come: but as soon as she is delivered of the child, she remembereth no more the anguish, for joy that a man is born into the world.";
  public static final String jon1621new = "JOH 16:21  A woman when she is giving birth [in travail] hath sorrow, because her hour is come: but as soon as she is delivered of the child, she remembereth no more the anguish, for joy that a man is born into the world.";
  public static final String jon1621newer = "JOH 16:21  A woman when she is giving birth [in travail] is sad [hath sorrow], because her hour is come: but as soon as she is delivered of the child, she remembereth no more the anguish, for joy that a man is born into the world.";
  public static final String jon1621newest = "JOH 16:21  A woman when she is giving birth [in travail] is sad [hath sorrow], because her hour is come: but as soon as she is delivered of the child, she remembereth no more the pain [anguish], for joy that a man is born into the world.";
  public static final String honorable = "Hez 2:22  honour to whom Honour to whom honour";
  public static final String honorablenew = "Hez 2:22  honor to whom Honor to whom honor";
  public static PassingItems pi;
  @BeforeAll
  static void setUpBeforeClass() throws Exception {
  }

  @AfterAll
  static void tearDownAfterClass() throws Exception {
  }

  @BeforeEach
  void setUp() throws Exception {
  }

  @AfterEach
  void tearDown() throws Exception {
  }

  @Test
  void testFindWordIndex() {
    // Test finding word1 - should find it at position 0
    int index1 = WordUpgradeUtils.findWordIndex(aLine, "word1");
    assertEquals(0, index1, "word1 should be found at index 0");
    
    // Test finding word3 - should find it at position 12
    int index3 = WordUpgradeUtils.findWordIndex(aLine, "word3");
    assertEquals(27, index3, "word3 should be found at index 27");
    
    // Test finding word4 - should not find it
    int index4 = WordUpgradeUtils.findWordIndex(aLine, "word4");
    assertEquals(-1, index4, "word4 should not be found");
    
    // Test finding word5 - should not be found because it's enclosed in []
    // and word5a has an alphabetic character after word5
    int index5 = WordUpgradeUtils.findWordIndex(aLine, "word5");
    assertEquals(-1, index5, "word5 should not be found because it is enclosed in []");
  }

  @Test
  void testModernizeWord() {
    pi = new PassingItems(g11, "oldWord", "newWord");
    assertEquals("GEN 1:1", pi.bookChapVerse, "Strip chap and verse off the line.");
    assertEquals(g11.substring(9), pi.line, "Remaining line after strip chap and verse."); 
    assertEquals(g11, pi.getEditedLine(), "Original line with no changes.");
    pi = new PassingItems(mat12, "begat", "fathered");
    WordUpgradeUtils.modernizeWord(pi);
    assertTrue(pi.isDirty);
    assertEquals(mat12New, pi.getEditedLine(), "begat to fathered");
    pi = new PassingItems(jon1621, "in travail", "giving birth");
    WordUpgradeUtils.modernizeWord(pi);
    assertTrue(pi.isDirty);
    assertEquals(jon1621new, pi.getEditedLine(), "in travail to giving birth");
    pi.setWords("hath sorrow", "is sad");
    WordUpgradeUtils.modernizeWord(pi);
    assertTrue(pi.isDirty);
    assertEquals(jon1621newer, pi.getEditedLine(), "hath sorrow to is sad");
    pi.setWords("anguish", "pain");
    WordUpgradeUtils.modernizeWord(pi);
    assertTrue(pi.isDirty);
    assertEquals(jon1621newest, pi.getEditedLine(), "anguish to is pain");
    StringBuilder sb = WordUpgradeUtils.wordChanges.get("hath sorrow");
    assertNotNull(sb);
    assertEquals("JOH 16:21 ", sb.toString(), "John 16:21 changed");
    //assertEquals(2, WordUpgradeUtils.verseChanges.size(), "changed 3 verses");
 }

  @Test
  void testReplaceWord() {
    pi = new PassingItems(honorable, "honour", "honor");
    WordUpgradeUtils.replaceWord(pi);
    assertTrue(pi.isDirty);
    assertEquals(honorablenew, pi.getEditedLine(), "upgrading honour");
  }
}
