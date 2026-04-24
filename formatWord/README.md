#Format Bibles for Microsoft Word

This program takes a bible formatted as one file per book.  It reads each verse, adds footnotes, TOC entries, and book introductions, and formats the result into a template document.

## Template documents

Each template has the introduction, defines styles, headers, and footers.  

**21stCentKJB6x9.docx** formats as a 6 inch X 9 inch book.

**21stCentKJB8x11.docx** formats as an 8.5 inch X 11 inch book.  The Old Testament is long enough to have to be formatted using this template.  The entire King James Bible can be formatted as one book, but the **Verse** and **FootnoteText** styles must be set to 8 point type to fit Amazon's page limit.

That leaves the Introduction, Table of Contents, and explanation on the back of the title page in 12 point.  You can use **SetStyleFontSizesMain** to adjust all the font sizes if that seems better to you.

Font sizes of styles in a .docx can be set using **SetStyleFontSizesMain** in the adjacent fixFonts folder.  Changing sizes from list of styles and sizes is faster and more accurate than doing it by hand. 

**ListStylesSizesDocxMain** lists font sizes for each style where size is specified.  If a style gets its size from a parent style, it has no size.  This program produces a tab-separated list of lines which can be pasted into Notepad and then into Excel.

Once the data are in Excel, other columns can be added which set the sizes in the target document.  This can be written to a .csv file which is read by **SetStyleFontSizesMain**.  There is a sample in the **FontSizes** tab of **KJBWordUpdates.xlsx** which has the list of old-fashioned words for which help is provided. 

## Helpful Scripts

The spreadsheet of old-fashioned words has one row per old-fashioned word.  Each row may have an Only in: with a list of verses to which the change applies or a Not in: list of verses where the change is not applied.  If neither is listed, every verse is checked for the old-fashioned word. 

Find verses in the Bible and collect all references to them in one line for use with Only in:

```bash
grep ... | sed -E 's/^[^:]*:[[:space:]]*//; s/^(([^[:space:]]+[[:space:]]+){2}).*/\1/; s/[[:space:]]+$//' | paste -sd' ' -
```

This keeps text after the colon through the 2nd space, removes the trailing space, and joins with a single space.  This string can be pasted into an Only in: field of the spreadsheet. Define this string as an alias:

```bash
 alias sedRefs="sed -E 's/^[^:]*:[[:space:]]*//; s/^(([^[:space:]]+[[:space:]]+){2}).*/\1/; s/[[:space:]]+$//' | paste -sd' ' -"
 ```

This **grep** command isolates words by requiring a non-word space both before and after the word. Searching for a word will not find any other forms.

```bash
grep -P -i '\WBeelzebub\W' *
```

Combining the two commands:

```bash
grep -P -i '\WBeelzebub\W' * | sedRefs
```

produces **MAT 10:25 MAT 12:24 MAT 12:27 MAR 3:22 LUK 11:15 LUK 11:18 LUK 11:19**.

## Snippets

- **AddRTFParamarksToTextKJVMain**: Reads a text file containing Bible verses which include paragraph symbols and insert paragraph symbols in the matching verses in another file.  This file will then be split into one file per book with the chapter and verse numberings expected by the Bible formatting programs.
- **CompareElectronicBiblesMain**: Compare two folders of book files to verify that all file names and all 31,102 chapter and verse numbers line up across both sources. This makes sure that either source can be used when formatting a gentle bible.
- **DiscoverStylesMethodsMain**: Use reflection to list XWPFStyles methods to explore POI style APIs.
- **ListBookmarkReferencesMain**: List bookmarks in a docx and the paragraphs that reference them (excluding _Toc).
- **ListBookmarksDocxMain**: List bookmark names and the paragraphs where they are defined (excluding _Toc).
- **ListStylesDocMain**: List styles in a legacy .doc file using HWPF.
- **ListStylesDocxReflectionMain**: List styles in a docx with style ID, name, and type via reflection.
- **ParseRtfToRawTextMain**: Convert any RTF file to plain text with inline &lt;i&gt;, &lt;b&gt;, and &lt;u&gt; tags. Other formatting information is ignored.
- **TxtKJBTo66ChapterFilesMain**: Split a Bible text file into 66 book files in the format expected by **gentlerKJB** and **formatWord**.

