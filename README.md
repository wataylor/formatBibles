# The Goal of this Ministry

The King James Bible is hard to read because the English language has changed since the popular 1769 version was published. We have friends who *know* that the KJV is the true Word of God but read the NIV because they understand it.

The King James has had more than 24,000 changes in spelling and printing since 1611. For example, “love” was printed as “loue,” “us” as “vs,” and “ever” as “euer.” “S” was printed “f” in the same way the Declaration of Independence wrote of the “purfuit of happineff.” 

Such updates were needed as English spelling and fonts changed over time and do not change the Word of God.  This project carries on such changes in a way that fulfills Paul's charges to Timothy which gave local churches responsibility for preserving and spreading the Very Words of God.

The first program reads an Excel spreadsheet listing archaic words and more modern substitutions.  It then reads through an electronic version of the 1769 King James looking for archaic words.  The modern equivalent is inserted and the archaic word is surrounded by [] to help readers learn them over time.

The second program formats the raw text into a print-ready .docx file which can be edited as desired.

# Public git Repository

The source code is in a public git repository [formatBibles](git@github.com:wataylor/formatBibles.git).  The program also uses some .docx and .xlsx files which are stored in [The Gentle King James Project](https://drive.google.com/drive/folders/1hdicavzgwZQDg9vzcX0BI7uwMEMPS6D2?usp=drive_link) public folder.

# Building All System Components

The top-level **pom.xml** builds the Important Jars and the Support Jars.  

# The Important Jars in the Repo

This repo generates two executable .jar files: **gentlerKJB** and **formatWord**.  They're run with **java -jar <filename> <command args>**. I can never remember all the arguments, so each executable has a **+help** argument which documents the default argument values and exits.

By a strange coincidence, the default values are what I need to run the programs based on where I chose to put the input and output files. **+help** is a default value so running the program with no arguments explains the documents and exists. If you want it to run, use **-help** to override the **+help**.

## gentlerKJB

This program makes the King James easier to read by changing phrases like “great with child” to “about to give birth” and “with child” to “pregnant.” The original words are put in [] so that Matthew 1:18-19 became:

Now the birth of Jesus Christ was like this [on this wise]: When as his mother Mary was [espoused] {engaged} to Joseph, before they came together, she was found pregnant [with child] of the Holy Ghost.  Then Joseph her husband, being a just man, and not willing to make her a publick example, was minded to put her away secretly [privily].

These updates make no spiritual or doctrinal changes whatsoever. If you find any changes that affect doctrine, **PLEASE SAY SO!!** and we’ll fix it.

Word changes are driven by an Excel file in [this Google Drive folder](https://drive.google.com/drive/folders/1hdicavzgwZQDg9vzcX0BI7uwMEMPS6D2?usp=drive_link):

The .xlsx file has many sheets defining footnotes, important verses, and bible files names which are explained in [The Gentle King James Project](https://docs.google.com/document/d/1YroHALejw5bwmQqKdDp5l4PGsidNIA9zChnktP8deyY/edit?usp=sharing):

The 1769 King James books whose archaic words are replaced is in the subfolder **KJVText**.

A more recent set of book files is in the **kjv-bibleprotector-com** subfolder.  These files were derived from a text file which was downloaded from the Australian site [Bible Protector](https://bibleprotector.com/): Australia is in the British Commonwealth which gives this group close connections to the KJB custodians at Cambridge University.

## formatWord

This program reads the chapter files modified by **gentlerKJB** and formats them into a .docx file.  Each chapter gets a section with a one-column chapter heading and a two-column section for the verses in the chapter.

The words being changed are listed in the Excel file **KJVWordUpdates.xlsx**. **GentleKJB6x9.docx** is a template that defines styles, footers, table of contents, and the introduction for a 6 inch by 9 inch book. **GentleKJB8x11.docx** is for 8.5 x 11 inch books. The output file command-line argument is created by injecting the modified verses into the template at the proper place.

The spreadsheet also defines footnotes for specific words in specific verses on the **Footnotes** sheet and a list of verse comments which are added to the Table of Contents in the **TOCVerses** sheet.  These sheets are explained in more detail in **The Gentle King James Project** description in the folder.

The program adds a list of all verses that were changed with hyperlinks to the changed verses. That list is followed by a list for each archaic word with hyperlinks to the verses it changed. This should help verify all the changes.

The **snippets** folder has a number of utility programs for manipulating text files and .docx files.

# The Support Jars

**commonClasses** builds a .jar file containing command utilities for parsing command-line arguments and for reading Excel spreadsheets.

**fixFonts** generates an executable .jar file which sets fonts in a .docx file to a specified value.  This is useful because Microsoft Word hides font names in many places.  A hidden font may not be embeddable.  If you want to produce a .pdf file containing only embeddable fonts, choosing an embeddable font for all uses in the document is a good first step.

The current version updates fonts in these major locations:
- styles and document defaults (`styles.xml`, `stylesWithEffects.xml`, including latent/default font nodes)
- numbering (`numbering.xml` list level and run font settings)
- themes (`/word/theme/*.xml`, including major/minor latin/ea/cs and script fonts)
- main document runs (`document.xml`) and runs in headers, footers, footnotes, and endnotes
- comments, settings defaults, and `fontTable.xml`

Command line:
- `SetAllStylesToAFontMain [-verbose] "Font Name" "path\\to\\file.docx"`
- `-verbose` enables detailed diagnostics for troubleshooting unusual Word behavior.

If you save a .docx as a .pdf, Word sometimes includes the Arial font for reasons known only to Microsoft but refuses to embed it.  Your .pdf file will be rejected by recipients who insist on embedded fonts.  In that case, using the Microsoft ""Print as PDF** printer may help because it seems to use only embeddable fonts.

If that won't work because it doesn't support your paper size, try using **Save As PDF** but before hitting **Save**, click **Options** and check the **ISO 9005-1 compliant (PDF/A)** box.  ISO 9005-1 .pdf files must include embedded fonts.  This forces Word to include Airal as an embedded font.

# ✨ I'm A Geek

I’ve been earning a living in what is now called IT since **June 1964**.
- My first job out of college was working on software for the Poseidon Missile and with the Apollo software that took us to the moon. I *am* a rocket scientist, so when I say something isn't rocket science, I speak with authority.
- In the early 1970s, I developed computerized typesetting software for *The New York Times* so they no longer needed molten lead to cast lines of type for the paper.
- In 1988, the MIT Press published my book "What Every Engineer Should Know About Artificial Intelligence." I formatted it using an early version of Microsoft Word.
- I’m now exploring document automation as a hobby, building a modular library of reusable methods — my “spellbook.” It isn't rocket science, but it has hidden subtleties and complexities.

This project wouldn’t have been practical as an unpaid ministry without the help of GitHub CoPilot guiding me in stitching together the pieces: POI code patterns, JUnit test suites, and practical workflows.

## 📖 Document Spellbook

This project contains a collection of Java “spells” built on [Apache POI](https://poi.apache.org/) for automating the creation of Word `.docx` documents. Each spell encapsulates a discrete trick — inserting bookmarks, footnotes, index entries, hyperlinks, and more — with accompanying JUnit tests to ensure reliability.


### 🛠 Features

- **Bookmark spells**: Insert named bookmarks anywhere in a document.
- **Hyperlink spells**: Create cross references that jump to bookmarks when clicked in Word or in generated .html files.
- **Footnote spells**: Add footnotes programmatically.
- **Index entry spells**: Insert XE fields for building an index.
- **Split heading spells**: Style headings and following text correctly.
- **Inline italics tag spell**: Convert balanced `<i>...</i>` markers into true italic runs.
- **Drop Text spell**: Put a text box containing a string at the beginning of a paragraph for chapter numbers.
- **Cross‑reference appendix**: Auto-generate a paragraph of hyperlinks to all bookmarks.
- **JUnit 5 test suite**: Verifies XML structures so you don’t need to open Word every time.

## 🔧 Utility Functions

A later section describe helpers methods that support the core spells.
Examples include:
- Bookmark ID generators
- Hyperlink styling helpers
- Paragraph separators
- Common test assertions
- Java programs to list style names, bookmark names, and bookmark references in a .docx.

## 🚀 Getting Started

Dependencies:
```xml
<dependency>
  <groupId>org.apache.poi</groupId>
  <artifactId>poi-ooxml</artifactId>
  <version>5.2.5</version>
</dependency>
<dependency>
  <groupId>org.junit.jupiter</groupId>
  <artifactId>junit-jupiter</artifactId>
  <version>5.11.0</version>
  <scope>test</scope>
</dependency>
```

Compile with Java 8 or higher. Run tests with JUnit 5.

## 📜 Philosophy

This repository is both technical and personal:
- Technical, because it demonstrates how to automate Word documents with POI and lock down correctness with JUnit.
- Personal, because it reflects a lifetime in IT — from hot‑metal typesetting to XML automation — and the joy of continuing to learn and share.

---

## 📚 Usage Examples

### Add a Bookmark
```java
XWPFDocument doc = new XWPFDocument();
DocumentSpellbook.addBookmarkParagraph(doc, "Chapter1Start", "This is Chapter 1");
```

Creates a bookmark named `Chapter1Start` at the start of a paragraph.

---

### Add a Footnote
```java
XWPFParagraph para = doc.createParagraph();
para.createRun().setText("Text with footnote");
DocumentSpellbook.addFootnote(para, doc, "Footnote text");
```

Adds a footnote to the paragraph with the text `"Footnote text"`.

---

### Add a Hyperlink to a Bookmark
```java
// Target bookmark
DocumentSpellbook.addBookmarkParagraph(doc, "target", "This is the target");

// Source hyperlink
XWPFParagraph para = doc.createParagraph();
DocumentSpellbook.addHyperlinkToBookmark(para, "target", "source");
```

Creates a hyperlink labeled `"source"` that jumps to the bookmark `"target"`.

---

### Convert `<i>` Tags to Italic Runs
```java
XWPFParagraph para = doc.createParagraph();
para.createRun().setText("This is <i>important</i> text.");
WordDocxUtils.applyItalicTags(para, doc);
```

Removes the `<i>` and `</i>` markers from visible text and formats the enclosed phrase in italics.

---

### Generate a Cross‑Reference Appendix
```java
List<String[]> refs = List.of(
    new String[]{"Section1", "See Section 1"},
    new String[]{"Section2", "See Section 2"},
    new String[]{"Section3", "See Section 3"}
);

DocumentSpellbook.addCrossReferenceAppendix(doc, refs);
```

Builds a paragraph with hyperlinks to multiple bookmarks, separated by `|`.

---

### Run JUnit 5 Tests
```java
@Test
void testBookmarkInserted() {
    DocumentSpellbook.addBookmarkParagraph(doc, "Intro", "Introduction");
    assertTrue(doc.getParagraphs().get(0).getCTP()
        .getBookmarkStartList()
        .stream()
        .anyMatch(b -> "Intro".equals(b.getName())));
}
```

Verifies that the bookmark `"Intro"` was correctly inserted into the document.

## 🧪 Test Coverage

This project includes a growing suite of **JUnit 5 tests** to verify that each spell produces the correct XML structures inside the `.docx`. Tests ensure you don’t need to open Word every time — you can trust the automation.

Currently covered:

- **Bookmarks**
  - Verify bookmark names and IDs are correctly inserted.
- **Footnotes**
  - Confirm footnote text is stored in the document’s footnote collection.
- **Superscript paragraphs**
  - Assert that runs are styled with `VerticalAlign.SUPERSCRIPT`.
- **Index entries (XE fields)**
  - Check that the paragraph contains the correct `XE` field instruction.
- **Split headings**
  - Ensure heading styles (`Heading2`, `Normal`) are applied to the right paragraphs.
- **Inline italics tag conversion**
  - Verify balanced `<i>...</i>` markers are removed and converted to italic runs.
- **Hyperlinks to bookmarks**
  - Verify that hyperlinks contain the correct anchor and visible text.
- **Cross‑reference appendix**
  - Confirm multiple hyperlinks are generated in one paragraph, separated by delimiters.

---

### ✅ Philosophy of Testing
- **Unit tests**: Inspect the XML structures directly (bookmarks, hyperlinks, footnotes, styles). You must make sure the results look right in Word yourself.
- **Integration checks**: Occasionally open the document in Word to confirm rendering, but rely primarily on automated tests.
- **Confidence**: Every new spell gets a test, so the spellbook grows with a safety net.
