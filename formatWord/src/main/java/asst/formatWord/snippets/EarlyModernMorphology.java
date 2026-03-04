package asst.formatWord.snippets;

/** Determine the morphology of words based on King James English
 * word endings
 * @author Copilot
 * @since 2026 03
 */
public class EarlyModernMorphology {

  /** Morphology tags tell which type a word is
   * 
   */
  public static class MorphTag {
    public final String lemma;
    public final String tag;

    /** Create a tag of a type
     * @param lemma
     * @param tag
     */
    public MorphTag(String lemma, String tag) {
      this.lemma = lemma;
      this.tag = tag;
    }

    @Override
    public String toString() {
      return "MorphTag{lemma='" + lemma + "', tag='" + tag + "'}";
    }
  }

  // --- Lexical exceptions that look like endings but are not ---
  private static final String[] LEXICAL_EXCEPTIONS = {
      "froward", "noisome", "hireling", "changeling", "suckling", "wedlock"
  };
  // Words that do not need analysis
  private static final String[] SKIP_WORDS = {
      "about", "acts", "adversaries", "afraid", "after", "against",
      "ahasuerus", "amen", "amorites", "anger", "angels", "anointed",
      "another", "answer", "answered", "apostles", "appeared",
      "appointed", "ashamed", "ashes", "asked", "asses", "assembled",
      "atonement", "abner", "abominations", "abroad",
      "bands", "baptized", "been", "beast", "beasts", "begat", "behind",
      "behold", "believed", "beloved", "better", "between", "bless",
      "blessed", "blind", "blood", "bones", "border", "bound", "bowed",
      "bowls", "branches", "brass", "bread", "brethren", "broken", "brother",
      "brought", "build", "built", "burden", "buried", "burned",
      "burnt",
      "cast", "called", "cannot", "captives", "captains", "carried",
      "caught", "caused", "chains", "chamber", "chambers", "chariot",
      "chariots", "cherubims", "child", "children", "chosen", "christ",
      "churches", "cities", "clothes", "clothed", "cloud", "coast",
      "coasts", "comfort", "comforted", "command", "commanded",
      "commandment", "commandments", "commit", "committed", "conceived",
      "consider", "consumed", "could", "countries", "court", "covenant",
      "covered", "created", "cried", "crucified", "cubit", "cubits",
      "cursed",
      "damascus", "darkness", "daughter", "daughters", "david",
      "days", "dead", "declared", "deeds", "defiled", "delight",
      "deliver", "delivered", "depart", "departed", "destroyed",
      "devils", "died", "disciples", "divided", "dogs", "doors",
      "dried", "driven", "dwelling",
      "east", "ears", "eaten", "egypt", "egyptians", "eight", "either",
      "elders", "eleven", "enemies", "enter", "entered", "ephod",
      "escaped", "established", "even", "ever", "except", "eyes",
      "fast", "faces", "faint", "fallen", "families", "father",
      "fathers", "feed", "feast", "feared", "feet", "fenced",
      "field", "fields", "fight", "filled", "find", "first",
      "fled", "flocks", "flood", "followed", "food", "foot",
      "forbid", "forget", "forgiven", "former", "fought", "found",
      "fowls", "friend", "friends", "fruit", "fruits",
      "garden", "garment", "garments", "gates", "gather", "gathered",
      "generations", "gentiles", "ghost", "gift", "gifts",
      "gilead", "given", "glad", "gladness", "goats", "gods",
      "gold", "golden", "good", "goods", "grapes", "grass",
      "green", "great", "ground", "groves", "guard",
      "hand", "hands", "harlot", "harvest", "hated", "heed",
      "head", "heads", "healed", "heard", "heart", "hearts",
      "heathen", "heaven", "heavens", "height", "held", "herod",
      "hills", "hired", "host", "hold", "horns", "horses",
      "horsemen", "hosts", "houses", "hundred", "hurt", "husband",
      "husbands",
      "idols", "images", "increased", "indeed", "inhabited",
      "inhabitants", "inherit", "iniquities", "innocent",
      "instead",
      "james", "jehoshaphat", "jesus", "jews", "just", "judged",
      "judges", "judgment", "judgements", "justified",
      "kept", "kind", "kindness", "kings", "kingdoms", "knees",
      "lambs", "lamps", "land", "left", "levites", "lift", "lifted",
      "light", "linen", "lions", "lips", "loaves", "lord",
      "manner", "master", "matter", "meet", "measured", "messenger",
      "messengers", "midst", "might", "minister", "most",
      "moses", "mother", "mount", "mountains", "must", "multiplied",
      "naked", "nakedness", "names", "nations", "neither", "never",
      "nevertheless", "next", "night", "number", "numbered",
      "offer", "offered", "offerings", "officers", "ones", "open",
      "oppressed", "order", "other", "others", "over", "oxen",
      "part", "parts", "pass", "passed", "passover", "paths",
      "perfect", "persons", "peter", "pharisees", "philistines",
      "pieces", "pillars", "pitched", "places", "possess", "power",
      "prayer", "preached", "preacher", "precioius", "prepared",
      "present", "priest", "priests", "princes", "prophet",
      "prophets",
      "rest", "raiment", "rams", "rather", "read", "received",
      "redeemed", "redeemer", "reigned", "remember", "remembered",
      "remnant", "removed", "repaired", "repent", "repented",
      "respect", "returned", "reuben", "reward", "rewards", "riches",
      "right", "righteous", "river", "rivers", "round", "ruler",
      "rulers",
      "sabbaths", "sacrifices", "said", "saints", "scarlet", "scribes",
      "seed", "seen", "seat", "second", "secret", "send", "sent",
      "serpent", "served", "servant", "servants", "seven", "should",
      "shoulder", "shut", "sides", "sight", "silver", "sins",
      "singers", "sinned", "sister", "skins", "slaughter", "slept",
      "sold", "sons", "souls", "sound", "spirit", "stand", "statutes",
      "staves", "stones", "stood", "stranger", "strangers", "streets",
      "suburbs", "sweet", "sword",
      "taught", "tent", "tempted", "tents", "then", "that", "thanks",
      "themselves", "this", "things", "third", "thousand", "thousands",
      "times","together", "told", "toward", "trees", "tribes",
      "trust", "turned",
      "under", "understand", "upright", "upward",
      "vessels", "villages",
      "wait", "walls", "washed", "water", "waters", "ways", "weight",
      "went", "when", "what", "wicked", "wilderness", "wind",
      "wings", "without", "witness", "wives", "women", "wood", "word",
      "words", "works", "world", "worshipped", "would", "written",
      "years", "yourselves",
  };

  /** Determine the classification of a KJB word based on its ending. 
   * @param word
   * @return a Morphology Tag or null
   */
  public static MorphTag classifyEarlyModernEnding(String word) {
    if ((word == null) || (word.length() <= 3)) return null;
    String w = word.toLowerCase();

    // 1. Protect lexical exceptions
    for (String ex : LEXICAL_EXCEPTIONS) {
      if (w.equals(ex)) {
	return new MorphTag(word, "lexical.exception");
      }
    }

    // 1.1 remote obvious words
    for (String ex : SKIP_WORDS) {
      if (w.equals(ex)) {
	return null;
      }
    }

    // 2. -eth → 3rd person singular present
    if (w.endsWith("eth") && w.length() > 3) {
      String lemma = w.substring(0, w.length() - 3);
      return new MorphTag(lemma + " " + w, "verb.present.3sg");
    }

    // 3. -est → 2nd person singular present
    if (w.endsWith("est") && w.length() > 3) {
      String lemma = w.substring(0, w.length() - 3);
      return new MorphTag(lemma + " " + w, "verb.present.2sg");
    }

    // 4. -st → 2nd person singular (short form)
    if (w.endsWith("st") && w.length() > 2) {
      String lemma = w.substring(0, w.length() - 2);
      return new MorphTag(lemma + " " + w, "verb.present.2sg");
    }

    // 5. -en → participle or archaic plural
    if (w.endsWith("en") && w.length() > 2) {
      String lemma = w.substring(0, w.length() - 2);
      // You can refine this with a strong-verb whitelist
      return new MorphTag(lemma + " " + w, "verb.participle.or.plural");
    }

    // 6. -ed / -d / -t → weak past tense
    if (w.endsWith("ed") && w.length() > 2) {
      String lemma = w.substring(0, w.length() - 2);
      return new MorphTag(lemma + " " + w, "verb.past");
    }
    if (w.endsWith("d") && w.length() > 1) {
      String lemma = w.substring(0, w.length() - 1);
      return new MorphTag(lemma + " " + w, "verb.past");
    }
    if (w.endsWith("t") && w.length() > 1) {
      String lemma = w.substring(0, w.length() - 1);
      return new MorphTag(lemma + " " + w, "verb.past");
    }

    // 7. -er / -est → comparative / superlative adjectives
    if (w.endsWith("er") && w.length() > 2) {
      String lemma = w.substring(0, w.length() - 2);
      return new MorphTag(lemma + " " + w, "adj.comparative");
    }
    if (w.endsWith("est") && w.length() > 3) {
      String lemma = w.substring(0, w.length() - 3);
      return new MorphTag(lemma + " " + w, "adj.superlative");
    }

    // 8. -ward / -wards → directional adverb
    if (w.endsWith("ward")) {
      String lemma = w.substring(0, w.length() - 4);
      return new MorphTag(lemma + " " + w, "adv.direction");
    }
    if (w.endsWith("wards")) {
      String lemma = w.substring(0, w.length() - 5);
      return new MorphTag(lemma + " " + w, "adv.direction");
    }

    // 9. -wise → manner adverb
    if (w.endsWith("wise") && w.length() > 4) {
      String lemma = w.substring(0, w.length() - 4);
      return new MorphTag(lemma + " " + w, "adv.manner");
    }

    // 10. -some → qualitative adjective
    if (w.endsWith("some") && w.length() > 4) {
      String lemma = w.substring(0, w.length() - 4);
      return new MorphTag(lemma + " " + w, "adj.qualitative");
    }

    // 11. -ling → agentive noun
    if (w.endsWith("ling") && w.length() > 4) {
      String lemma = w.substring(0, w.length() - 4);
      return new MorphTag(lemma + " " + w, "noun.agentive");
    }

    // 12. Genitive -s (no apostrophe)
    if (w.endsWith("s") && w.length() > 1) {
      String lemma = w.substring(0, w.length() - 1);
      return new MorphTag(lemma + " " + w, "noun.genitive");
    }

    // 13. Default: no detectable archaic ending
    return null; //return new MorphTag(word, "unclassified");
  }
}  

