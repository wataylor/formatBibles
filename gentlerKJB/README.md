# gentlerKJB

Java 8 command-line tool that reads text files, searches a
spreadsheet for old-fashioned words, and adds helping words.

Build:

```bash
mvn clean install
```

Run:

Bash command to see the default command-line settings, add -help to turn off help and run the program:

```bash
java -jar target/gentlerKJB-0.0.1-SNAPSHOT.jar
```

## Preparing the Gentle Version

The program prepares Bible verses to be formatted into a Word document.  Edited chapter files are written into an output directory, and an explanation of the result into a file named **explanation.txt**.

These files are read and formatted by a companion program.
