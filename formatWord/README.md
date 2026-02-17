#Format Bibles

This program takes a bible formatted as one file per book.  It reads each verse, adds footnotes, TOC entries, and book introductions, and formats the result into a template document.

## Template documents

Each template has the introduction, defines styles, headers, and footers.  

**GentleKJB6x9.docx** formats as a 6 inch X 9 inch book.

**GentleKJB8x11.docx** formats as an 8.5 inch X 11 inch book.  The Old Testament is long enough to have to be formatted using this template.  The entire King James Bible can be formatted as one book, but it has to be in 8 point type to fit Amazon's page limit.

Font sizes of styles in a .docx can be set using **SetStyleFontSizesMain** in the adjacent fixFonts folder.  Changing sizes from list of styles and sizes is faster and more accurate than doing it by hand. 

**ListStylesSizesDocxMain** lists font sizes for each style where size is specified.  If a style gets its size from a parent style, it has no size.  This program produces a tab-separated list of lines which can be pasted into Notepad and then into Excel.

Once the data are in Excel, other columns can be added which set the sizes in the target document.  This can be written to a .csv file which is read by **SetStyleFontSizesMain**.  There is a sample in the **FontSizes** tab of **KJBWordUpdates.xlsx** which has the list of old-fashioned words for which help is provided. 

## Helpful Scripts

The spreadsheet of old-fashioned words has one row per old-fashioned word.  Each row may have an Only in: with a list of verses to which the change applies or a Not in: list of verses where the change is not applied.  If neither is listed, every verse is checked for the old-fashioned word. 

Find verses in the Bible and collect all references to them in one line for use with Only in:

```bash
grep ... | sed -E 's/^[^:]*:[[:space:]]*//; s/^(([^[:space:]]+[[:space:]]+){2}).*/\1/; s/[[:space:]]+$//' | paste -sd' ' -
```

This keeps text after the colon through the 2nd space, removes the trailing space, and joins with a single space.  This string can be pasted into an Only in: field of the spreadsheet.