package asst.formatWord.snippets;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.OutputStreamWriter;
import java.util.Map;
import java.util.TreeMap;

public class WordMorphologyMain {

  public static void main(String[] args) {
    BufferedReader reader = null;
    BufferedWriter writer;
    String inputFile = null;
    String line = null;
    String targetLine = null;
    String inputPath = "C:\\Sync\\Biblical\\asciiBible\\";
    String outputFile = "C:\\Temp\\KJB\\Morphisms.txt";
    String chapVerse = null;
    String woids[];
    int morphs = 0;
    int morphedWords = 0;
    Map<String, StringBuilder> lemmaMap = new TreeMap<String, StringBuilder>();

    try {
      // Open the new output file
      writer = new BufferedWriter(
	  new OutputStreamWriter(
	      new FileOutputStream(outputFile),
	      "UTF-8"));
      for (String f : TxtKJBTo66ChapterFilesMain.BOOK_FILE_NAMES) {
	if (f.length() < 4) { continue; }
	inputFile = inputPath + f;
	reader = new BufferedReader(new FileReader(inputFile));
	while ((line = reader.readLine()) != null) {
	  int ix = line.indexOf("  ");
	  chapVerse = line.substring(0, ix);
	  targetLine = line.substring(ix+2)
	      .replaceAll("[^A-Za-z0-9\\s]", " ")
	      .replaceAll("[^A-Za-z0-9\\s]", " ")
	      .trim().toLowerCase();
	  woids = targetLine.split(" ");
	  for (String woid : woids) {
	    EarlyModernMorphology.MorphTag tag =
		EarlyModernMorphology.classifyEarlyModernEnding(woid);
	    if (tag == null) { continue; }
	    StringBuilder sb = lemmaMap.get(tag.lemma);
	    if (sb == null) {
	      sb = new StringBuilder(tag.tag + " ");
	      lemmaMap.put(tag.lemma, sb);
	      morphs++;
	    }
	    sb.append(chapVerse + "_");
	    morphedWords++;
	  }
	}
      }
      /* Dump the woids*/
      String msg = "Found " + morphedWords 
	  + " morphological words in "
	  + morphs + " categories.";
      System.out.println(msg);
      writer.write(msg);
      writer.newLine();
      for (Map.Entry<String, StringBuilder> e : lemmaMap.entrySet()) {
	msg = e.getKey() + " " + e.getValue().toString();
	System.out.println(msg);
	writer.write(msg);
	writer.newLine();
	writer.newLine();	
      }
      writer.close();
    } catch (Exception e) {
      System.out.println("File " + inputFile + " line " + line);
      e.printStackTrace();
    } 
  }
}

