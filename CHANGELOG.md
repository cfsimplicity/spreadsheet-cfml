## 4.6.0 - 27 February 2025

- Enhancements
	- \#394 Add hideSheet() and unhideSheet()

## 4.5.0 - 19 February 2025

- Enhancements
	- \#392 Add processLargeFile() method to process large XLSX files without reading the entire data into memory

## 4.4.0 - 8 February 2025

- Enhancements
	- \#297 readLargeFile() now works on ACF

## 4.3.1 - 29 January 2025

Fixes:
 - \#389 Faulty timezone check causes performance issues when handling date/time values in Lucee
 - \#390 Improve performance around date value handling

## 4.3.0 - 18 January 2025

- Enhancements
	-	\#383 Support java class loading via a dynamic path and improve configurability
	- \#384 Add java version to getEnvironment() details
	- **Experimental and only partial** support for BoxLang

- Maintenance
	- \#385 Upgrade POI to 5.4.0
	- \#386 Upgrade commons-csv to 1.12.0
	- \#388 Upgrade excel-streaming-reader to 5.0.3

## 4.2.1 - 23 August 2024

- Fix: \#377 AddRows() should not convert empty values to zeros when they are in a numeric typed query column

## 4.2.0 - 6 August 2024

- Enhancements
	- \#373 Allow datatype to be specified with addColumn()
	- \#374 Rename setCellValue() "type" argument to "datatype" for consistency

- Maintenance
	-\#375 Upgrade excel-streaming-reader to 5.0.2

## 4.1.1 - 9 July 2024

- Maintenance
	-\#372 Upgrade POI to 5.3.0, commons-csv to 1.11.0 and excel-streaming-reader to 4.4.0

## 4.1.0 - 22 June 2024

- Enhancements
	- \#369 Add moveSheet()
	- \#370 Add sheet position to sheetInfo() properties

- Maintenance
	-\#371 Upgrade excel-streaming-reader to 4.3.1

## 4.0.0 - 6 March 2024

- Breaking changes
	- \#325 Drop support for ACF2016
	- \#359 Library should default to returning cached cell formula values instead of always recalculating

- Enhancements
	- \#358 Allow control of whether to return cached or freshly calculated formula values
	- \#354 Add getCellAddress() to return a cell's alphanumeric reference
	- \#364 Support integer range validation
	- \#356 Support date range validation
	- \#363 Add option to allow New() to default to creating an XLSX (XML) spreadsheet object

## 3.12.0 - 27 November 2023

- Enhancements
	-\#346 Add writeCsv()

- Fixes
	-\#347/\#348 Avoid Perl/Java regex engine incompatibilities

- Maintenance
	-\#349 Upgrade POI to 5.2.5
	-\#350 Upgrade excel-streaming-reader to 4.2.1

## 3.11.1 - 15 November 2023

- \#345 readCsv(): Commons CSV boolean options should default to true 

## 3.11.0 - 14 November 2023

- Enhancements
	- \#343 Improve csvToQuery() performance
	- \#340 Add readCsv() for large file/advanced csv processing
	- \#336 Add parallelization option to queryToCsv()

- Fixes
  - \#339 csvToQuery() ignores trim setting when reading from file

- Maintenance
  - \#344 Upgrade excel-streaming-reader to 4.2.0

## 3.10.0 - 29 September 2023

- Maintenance
  - \#334 Upgrade POI to 5.2.4 and excel-streaming-reader to 4.1.2
  - \#324 Upgrade Commons CSV to 1.10.0

- Enhancements
	- \#331 Allow custom date formats to be set on an instance post-init()
	- \#328 Improve exception message when setRowHeight() is used with a non-existent row

## 3.9.0 - 2 May 2023

- Enhancements
	- \#321 Add createJavaObject() to support creating POI and other objects from the bundled jars
	- \#323 Add basic support for conditional formatting

- Fixes
  - \#322 Some values are not converted to hyperlinks when using URL datatype

## 3.8.1 - 8 March 2023

- \#320 Avoid cellStyle duplication when formatting cells from a struct over multiple calls

## 3.8.0 - 3 March 2023

- Enhancements
	- \#316 Support new override data types: url, email and file to auto-create hyperlinks when adding data
	- \#318 Allow the format argument of formatting methods to be a re-usable cellStyle object instead of a struct

- Fixes
  - \#317 setCellHyperlink() should re-use a single cell style over multiple calls by default

- Maintenance
  - \#319 Upgrade excel-streaming-reader to 4.0.5

## 3.7.4 - 18 December 2022

- \#313 queryToCsv() should not treat date strings in the data as date objects to be formatted

## 3.7.3 - 22 November 2022

- \#312 Move hosted files to forgebox, github is blocking download access preventing box install

## 3.7.2 - 17 November 2022

- \#311 Regression: csvToQuery() no longer works when file path is VFS

## 3.7.0 - 14 November 2022

- Enhancements
	- \#304 Improve csvToQuery() performance
	- \#308 Add option to return visible or raw value from getCellValue()
	- \#309 Add option to return visible values from read() and readLargeFile()

- Fixes
	- \#306 Chainable read() method should return the data if format specified
	- \#307 getCellFormat() throws error if XLSX cellFont has no colour value

- \#310 Upgrade excel-streaming-reader to 4.0.4

## 3.6.1 - 14 October 2022

- \#301 DataValidation has incorrect values if pulled from a sheet name which includes a space

## 3.6.0 - 12 October 2022

- Enhancements
	- \#300 Add support for creating DataValidation dropdowns

## 3.5.1 - 19 September 2022

- \#298 Upgrade POI to 5.2.3
- \#299 Upgrade excel-streaming-reader to 4.0.2

## 3.5.0 - 8 August 2022

- Enhancements
	- \#291 Speed improvement for getAllSheetFormulas()
	- \#293 Add includeHiddenRows option to read()
	- \#296 Add readLargeFile()

## 3.4.4 - 25 March 2022

- \#289 Prevent one-off OSGi bundle errors when the bundle version changes
- \#290 Upgrade POI to 5.2.2

## 3.4.3 - 8 March 2022

- \#288 sheetInfo() should default to the currently active sheet, not the first

## 3.4.1 - 7 March 2022

- \#287 Upgrade POI to 5.2.1

## 3.4.0 - 4 March 2022

- Enhancements
	- \#286 Add sheetInfo() to return metadata for a specific sheet within a workbook 

- Fixes
	- \#283 Mismatched system/Lucee timezones causes read() to offset date values
	- \#285 read() should use the first visible sheet in the workbook if no sheet is specified, ignoring hidden sheets

## 3.3.0 - 15 January 2022

- \#282 Upgrade POI to 5.2.0

## 3.2.5 - 30 December 2021

- Security Update
	- \#279 Upgrade log4j to 2.17.1

## 3.2.4 - 28 December 2021

- \#280 Fix: formatRows() errors if range specified is a single row

## 3.2.3 - 18 December 2021

- Security Update
	- \#279 Upgrade log4j to 2.17.0

## 3.2.2 - 17 December 2021

- Fixes
	- \#277 Date format initialization doesn't work in Lucee with full null support
	- \#278 Adding header/footer images throws error with null support enabled

## 3.2.1 - 15 December 2021

- Security Update
	- \#276 Upgrade log4j to 2.16.0

## 3.2.0 - 11 December 2021

- Security Update
	- \#273 Upgrade log4j to 2.15.0

- Enhancements
	- \#268 Allow row/column ranges specified for read(), deleteColumns(), deleteRows(), formatColumns() and formatRows() to be open-ended
	- \#274 Allow flushOsgiBundle() to flush a specified version

- Fixes
	- \#272 Read() not importing trailing empty columns

## 3.1.0 - 2 November 2021

- Enhancements
	- \#266 Upgrade POI to 5.1.0
	- \#261 Update commons-csv to 1.9.0

- Fixes
	- \#260 Chainable getCellComments() should return an array
	- \#262 Handle incorrect date value setting when Lucee timezone does not match system timezone
	- \#253 Fix autoSizeColumns not being applied to all columns when adding rows to streaming xlsx workbooks
	- \#263 Fix read() error if headerRow is specified and spreadsheet is empty
	- \#265 read( format="query" ) should auto-generate column names where too few column names are specified

## 3.0.0 - 17 September 2021

- Breaking changes
	- Rename project "spreadsheet-cfml"

- Enhancements
	- \#258 Add support for chainable operations on a workbook
	- \#257 Add getLastRowNumber()
	- \#254 Allow chaining of methods returning void

- Fixes
	- \#259 Fix error with addSplitPane
	- \#256 Improve performance of autoSizeColumns on addRows() when data is an array

## 2.21.0 - 29 July 2021

- Enhancements
	- \#251 Add makeColumnNamesSafe option to read() and csvToQuery()
	- \#252 Allow WriteToCSV to exclude the workbook's header row

## 2.20.0 - 18 June 2021

- Enhancements
	- \#248 Add autoSizeColumns option to `workbookFromQuery()`
	- \#246 Add `isCsvOrTextFile` support for .tsv files

- Fixes
	- \#250 Multi-cell formatting methods throw invalid arguments exception if overwriteCurrentStyle is set to false

## 2.19.0 - 14 May 2021

- Enhancements
	- \#240 Add setActiveCell()
	- \#238 Add setRecalculateFormulasOnNextOpen()

- Fixes
	- \#243 Handle null returned from getXSSFColor()
	- \#242 Ensure HeaderImageVML java is compiled for Java 1.8
	-	\#237 AddColumn() with startColumn and insert=true replaces the existing column instead of inserting after it

## 2.18.2 - 16 April 2021

- \#236 Regression: getCellFormula should not error if cell is specified but doesn't exist

## 2.18.1 - 15 April 2021

- Enhancements
	- \#231 Add setCellHyperLink() and getCellHyperLink()
	- \#101 Add setHeaderImage() and setFooterImage()

- Fixes
	- \#235 Fix missing semi-colon in setActiveSheetNameOrNumber()
	- \#233 Using "overwriteCurrentStyle=false" and a pre-built cellStyle with formatting functions causes the cellStyle to be ignored
	- \#232 Using "overwriteCurrentStyle=false" with formatting functions causes default cell style to be changed

## 2.17.0 - 12 March 2021

- Enhancements
	- \#229 Allow read() and csvToQuery() to accept a default queryColumnType
	- \#228 Allow read() to accept columnNames as an array as well as a list
	- \#227 Allow column names to be specified when using csvToQuery()
	- \#226 Allow query column types to be specified or auto-detected when using csvToQuery()

## 2.16.0 - 9 March 2021

- Enhancements
	- \#225 Allow query column types to be specified or auto-detected when reading a spreadsheet into a query

## 2.15.0 - 3 March 2021

- Enhancements
	- \#221 Add queryToCsv() and writeToCsv()
	- \#218 Add getPOIVersion()

- Fixes
	- \#219 dumpPathToClass() doesn't include file path in OSGi
	- \#222 \#223 Column/header names generated by csvToQuery() are in upper case in ACF

## 2.14.0 - 20 January 2021

- \#216 Upgrade POI to 5.0.0
- \#217 In Lucee use OSGi to load java classes instead of JavaLoader 

## 2.13.0 - 11 December 2020

- Enhancements
	- \#212 Support CSV custom delimiters in read() and downloadCsvFromFile()

## 2.12.2 - 12 November 2020

- Fixes
	- \#209 ACF2021: Class not found: org.apache.commons.io.output.ByteArrayOutputStream
	- \#208 ACF2021: Issue using includes in Testbox suites

## 2.12.1 - 30 October 2020

- \#206 Fix typo in downloadFileFromQuery()

## 2.12.0 - 22 October 2020

- Enhancements
	- \#204 Add public createCellStyle() method

- Fixes
	- \#205 Fix and improve handling of tab delimited data handling in csvToQuery()

## 2.11.1 - 9 September 2020

- \#202 Bugfix for isHex regex

## 2.11.0 - 4 September 2020

- Enhancements
	- \#201 Prevent ACF from treating "9a" or "9p" as a date/time value
	- \#200 When adding rows allow default data types to be overridden
	- \#199 Allow rows generated from queries to ignore the query column data types
	- \#198 Allow addColumn() to take data as an array
	- \#197 Add support for valid 6 character hexadecimal colors

## 2.10.0 - 14 April 2020

- Enhancements
	- \#190 Add getCellComments() as alias for getCellComment() with no row/column specified
	- \#188 Allow addAutoFilter to accept a row number instead of a cell range
	- \#187 Allow addAutoFilter to default to the first row

- Fixes
	- \#194 setCellComment() with underline throws exception on ACF
	- \#193 Prevent setCellComment() throwing an exception on XLSX when unsupported styles are set
	- \#191 setCellComment() throws exception on XLSX
	- \#189 Row and column number values missing from getCellComment() structs when all returned from sheet 
	- \#109 Write encryption doesn't work on ACF

## 2.9.0 - 3 April 2020

- Enhancements
	- \#186 Add option to formatting methods to preserve existing cell styles
	- \#184 Add support for DATETIME and DATETIME2 (MSSQL) database column types

- Fixes
	- \#185 Time only values do not respect custom TIME format specifying fractions of a second

## 2.8.0 - 23 March 2020

- Enhancements
	- \#181 Add "INT" to query column formats cast as numeric
	- \#179 Provide a list of all predefined colours available to formatting methods

- Fixes
	- \#182 addInfo() not working with Streaming XLSX 
	- \#178 Color index lookup is using a deprecated enum class

## 2.7.0 - 17 February 2020

- \#175 Upgrade POI to 4.1.2
- \#176 Upgrade Apache Commons CSV to 1.8

## 2.6.0 - 28 November 2019

- Enhancements
	- \#174 Add `getColumnWidth()` and `getColumnWidthInPixels()`

- Fixes
	- \#173 Specifying a custom DATETIME date format mask seems to have no effect
	-	\#172 In ACF query column case and order is not preserved
	- \#171 Using autoSizeColumns with a Streaming XLSX workbook causes an exception

## 2.5.0 - 21 November 2019

- \#170 Upgrade POI to 4.1.1
- \#169 Improve handling of clearly non-date values which Lucee will parse as dates far in the future

## 2.4.0 - 11 August 2019

- \#168 Allow the active sheet's "fit to page" print options to be controlled

## 2.3.0 - 9 August 2019

- \#167 Add support for setting sheet print margins

## 2.2.1 - 10 July 2019

- \#164 Upgrade Apache Commons CSV to 1.6
- \#166 Bug fix: autoSizeColumn - key [columnIndex] doesn't exist in argument scope

## 2.2.0 - 12 April 2019

- \#163 Upgrade POI to 4.1.0
- \#162 Support decryption of encrypted binary (XLS) spreadsheets. Add support for decryption in ACF.
- DEPRECATED: `engineSupportsEncryption` environment key. Use `engineSupportsWriteEncryption`

## 2.1.1 - 29 March 2019

- \#160 `includeQueryColumnNames` in `addRows()` produces invalid xlsx

## 2.1.0 - 15 January 2019

- \#159 Support array data argument for `addRow()` and `addRows()`

## 2.0.2 - 6 January 2019

- \#157 `addRows()` doesn't apply column offset to header row when using `includeQueryColumnNames`

## 2.0.1 - 22 December 2018

- \#156 Bug in `setActiveSheet()`.

## 2.0.0 - 22 December 2018

- Breaking changes
	- \#142 Upgrade POI to 4.0.1 which requires Java 8+
	- By default, Lucee 5 now uses JavaLoader instead of `CreateObject`
	- \#148 Remove the `engineSupportsDynamicClassLoading` variable completely, since it is meaningless
	- Remove Lucee 4.5 and ACF11 support: Lucee 5 and ACF2016 are the minimum supported versions

- Enhancements
	- \#155 Add support for the SXSSF streaming XML format for writing large files
  - \#136 Upgrade Apache Commons CSV to version 1.5
  - Improve `write()` outputstream locking.
  - Add `dumpPathToClass()` diagnostic tool
  - Separate encryption/decryption components no longer needed with POI 4

- Fixes
	- Various fixes to support POI 4.x
	- \#150 Rewrite xlsx encryption to ensure the encrypted stream is closed
	- Use array `append()` BIF instead of java `add()`
	- Fix failing `setCellValue()` test on ACF2016+
	- \#154 Using RGB triplet as a colour format with XLSX not working in ACF

## 1.7.3 - 7 November 2018

- Fixes
	- \#153 Handling of hidden columns fails in ACF2016
	- \#152 Testbox should be specified as a CommandBox installation dev dependency

## 1.7.2 - 14 May 2018

- Fixes
	- \#139 Cell type auto-detection throws error if boolean value is blank or null

## 1.7.1 - 13 May 2018

- Fixes
	- \#138 Reading a spreadsheet with column header names containing commas into a query results in too many columns

## 1.7.0 - 28 September 2017

- \#134 Upgrade POI to 3.17

## 1.6.1 - 7 September 2017

- Fixes
	- \#130 JavaLoader should not need `loadColdFusionClassPath` setting (add commons-codec jar to lib)

## 1.6.0 - 5 September 2017

- Enhancements:
	- \#129 Add `getRowCount()`

## 1.5.1 - 7 August 2017

- Fixes
	- \#128 Adding date query cells with blank values causes error

## 1.5.0 - 2 June 2017

- Enhancements
	- \#125 Add `addAutofilter()`
	- \#126 Add `addPageBreaks()`
	- \#127 Add `addPrintGridlines()` and `removePrintGuidelines()`

## 1.4.1 - 16 May 2017

- Fixes
	- \#124 Handle "indent" format values greater than 15 in xls
	- \#123 Underline value returned by `getCellFormat()` should be a descriptive string
	- Fix unreturned cellStyle when invalid underline format used.

## 1.4.0 - 15 May 2017

- Enhancements
	- \#119 Add `getCellFormat()` and formatting tests
	- \#104 Add `quoteprefixed` to cell formatting options
	- \#121 Support "double", "single accounting" and "double accounting" underline formats
	- \#122 Add `hideRow()`, `showRow()` and `isRowHidden()`
	- Add `isColumnHidden()` and tests for `hideColumn()` and `showColumn()`
- Fixes
	- \#120 Setting underline format on ACF fails

## 1.3.0 - 20 April 2017

- \#118 Upgrade POI to 3.16

## 1.2.0 - 18 April 2017

- Enhancements:
	- Rewrite cell data type handling
	- \#112 Allow data type to be specified when using `setCellValue()`
	- Add `getCellType()`
- Fixes:
	- \#115: Don't auto-detect any incoming values as boolean: allow them to default to strings or numbers, unless a source query column type or data type parameter has set them as bit/boolean
	- \#116: Prevent certain definitely non-date values being detected as dates
	- \#117: Allow `csvToQuery()` to be called with positional arguments

## 1.1.0 - 11 April 2017

- Enhancement: \#110 Support populating a spreadsheet from CSV data:
  - Add `workbookFromCsv()`
  - Add public method `csvToQuery()` for convenience

## 1.0.0 - 7 April 2017

- Enhancement: \#80 Provide option to use POI jars in the java class path instead of via JavaLoader
- Enhancement: \#107 Remove the dependency on JavaLoader for Lucee 5 and load POI jars directly
- Enhancement: \#108 Officially support ACF11+
- Enhancement: Add `getEnvironment()` method to return current environment details/settings

## 0.11.1 - 24 March 2017

- Update bundled JavaLoader to 1.2

## 0.11.0 - 9 March 2017

- Fix: \#102 Fillpattern formatting not working.
- Fix: \#103 Replace deprecated `cellstyle` setters.
- Enhancement: \#106 Allow a file path to be passed to the `info()` method instead of a workbook object

## 0.10.2 - 16 January 2017

- Fix: \#100 Replace deprecated 'boldweight' methods and constants with `getBold()`/`setBold()`.

## 0.10.1 - 16 January 2017

- Fix: Tweak argument handling of `setSheetPrintOrientation()` to catch errors.

## 0.10.0 - 15 January 2017

- Enhancement: \#99 Add `setSheetPrintOrientation()`.

## 0.9.4 - 21 December 2016

- Enhancement: \#98 Allow cell format color values to be specified in RBG triplets
- Fix: Refactor `buildCellStyle()` to improve performance and fix boolean issue with "bold".

## 0.8.4 - 01 December 2016

- Enhancement: \#97 Add `newXlsx()` and `newXls()` as aliases for `new( xmlFormat=true/false )`
- Fix: \#96 Fix error when adding image to xlsx (xml format) spreadsheet

## 0.8.3 - 22 November 2016

- Fix \#95 Change to `readBinary()` in 0.8.2 causes MS Excel to crash

## 0.8.2 - 17 October 2016
- Enhancement: Improve performance of `readBinary()` by using java ByteBuffer.
- Fix: Update JavaLoader to include post 1.1 release patches to fix \#94

## 0.8.1 - 29 September 2016
- Enhancement: \#92 Catch formula errors when reading

## 0.8.0 - 26 September 2016
- Enhancements:
	- \#90 Upgrade POI to 3.15
	- \#91 Allow an existing JavaLoader installation to be used instead of the bundled one.

## 0.7.9 - 4 August 2016
- Fix \#88 When reading a file, warn if the Excel format is too old for POI.

## 0.7.8 - 25 July 2016
- Fix \#87 Invalid color after moving from Lucee 4.5.3 to Lucee 5.0

## 0.7.7 - 14 July 2016
- Fix \#86 Zeros are being interpreted as boolean false by `addRow()` and other methods.

## 0.7.6 - 14 July 2016
- Fix \#85 `AddRow()` causing "maximum number of cell styles was exceeded" error when inserting large number of rows including dates.

## 0.7.5 - 13 July 2016
- Fix \#84 `formatColumn()` fails when workbook contains more than 4000 rows

## 0.7.4 - 1 July 2016
- Fix `isSpreadsheetFile()` not working in ACF for non-spreadsheet files.

## 0.7.3 - 1 July 2016
- Enhancements:
	- \#83 Add `isSpreadsheetFile()` and `isSpreadsheetObject()`

## 0.7.2 - 11 June 2016
- More ACF compatibility fixes
	- Drop use of `this` scope for internal method calls
	- Drop unnecessary `ExpandPath()` when getting POI jar paths for JavaLoader
	- Another missing colon at EOL

## 0.7.1 - 18 May 2016
- Updates to test suite for case-sensitive filesystems

## 0.7.0 - 17 May 2016
- More ACF compatibility fixes
	- Move all private methods from includes to within the body of Spreadsheet.cfc
	- Another missing colon at EOL
	- Use compatible script syntax for downloads
	- Variable name being used twice for different purposes

## 0.6.1 - 28 April 2016
- Fixes (preventing use with ACF):
 - Missing colons at EOL
 - Throw attribute typo

## 0.6.0 - 9 March 2016
- Enhancements:
	- \#76 Upgrade POI to 3.14
	- \#77 Add `getColumnCount()`

## 0.5.11 - 7 January 2016
- Better exception message when adding too many rows to a binary spreadsheet.

## 0.5.10 - 30 December 2015
- Better exception message when `read()` `src` file is not a spreadsheet.
- Make final closing of java streams dependent on existence of stream variable to prevent embedded exceptions.

## 0.5.9 - 28 December 2015
- Enhancements:
	- \#70  Support reading encrypted XML files (only) using supplied password
	- Better file handling when encrypting files

## 0.5.8 - 23 December 2015
- Enhancements:
	- \#73  Allow encryption algorithm to be specified when writing password protected workbooks

## 0.5.7 - 23 December 2015
- Fixes:
	-	\#71 and \#72	Adjust tests to support TestBox 2.2

## 0.5.6 - 15 December 2015
- Enhancements:
	- \#69  Add explicit `setReadOnly()` method for binary workbooks (only)
- Fixes:
	-	\#68	Supplying a password to `write()` provides whole file encryption for XML spreadsheet files (only)

## 0.5.5 - 3 December 2015
- Enhancements:
	- \#66 Add `setRepeatingColumns()` and `setRepeatingRows()`

## 0.5.4 - 12 November 2015
- Enhancements:
	- \#65 Upgrade POI to 3.13.
	- \#64 Add `download()` function for an existing workbook object.
	- \#63 Add `includeQueryColumnNames` option to `AddRows()`.

## 0.5.3 - 5 September 2015
- Improve performance of `read()` by using native Java concatenation instead of arrays which are slow in Lucee.

## 0.5.2 - 21 August 2015
- REMOVED:
 - \#61 Support for font family and size with `read()` and `includeRichTextFormatting` when different from the cell's base font. Better to be consistent and not support these attributes anywhere so the expectation is clear.

## 0.5.1 - 21 August 2015
- Enhancements:
 - \#61 Support font family and size with `read()` and `includeRichTextFormatting` when different from the cell's base font
- Bug fixes:
 - \#60 `includeRichTextFormatting` option in `read()` results in empty span style if format not supported

## 0.5.0 - 20 August 2015
- Enhancements:
 - \#57 Add `includeRichTextFormatting` option to `read()`

## 0.4.9 - 29 July 2015
- Enhancements:
 - \#56 Add extra argument to `read()` to allow excluding hidden columns
 - \#58 Add `hideColumn()` and `showColumn()`

## 0.4.8 - 8 June 2015
- Enhancements:
 - \#52 Add csv format support to `read()`
 - \#55 Allow csv file to be downloaded from a spreadsheet file

## 0.4.6 - 6 June 2015
- Enhancements:
 - \#43 Add html format support to `read()`.
 - \#54 Allow default date formats to be overridden
- Bug fixes:
 - \#53 Fix incorrect formatter reference when evaluating formula cells.

## 0.4.5 - 3 June 2015
- Enhancements:
 - \#44 Support reading specified row or column ranges
 - \#45 Support being able to specify the column names when reading a spreadsheet from file

## 0.4.4 - 31 May 2015
-	Bug fix:
 	- \#51 Empty cells are skipped when reading a spreadsheet into a query.

## 0.4.3 - 29 May 2015
- Upgrade POI to 3.12

## 0.4.2 - 29 May 2015
- Enhancements:
	- \#47 Add `fillMergedCellsWithVisibleValue` option to `read()`
	- \#48 Add `setCellRangeValue()`
	- \#49 Add `emptyInvisibleCells` option to `mergeCells()`
- Bug fixes:
	- Fix read() includeBlankRows=false option only suppressing null rows and not empty ones
	- Missing var declarations

## 0.4.1 - 10 March 2015
- Bug fix:
	- POI Loader server variable name should be unique to the current library path

## 0.4.0 - 25 February 2015
- Breaking changes
	- Use "freeze" instead of "split" for argument names of addFreezePane
	
## 0.3.0 - 24 February 2015
- Breaking changes
	- \#27 Drop `deleteSheet[Number]()` in favour of `removeSheet[Number]()`
- Bug fixes:
	- \#25 Font values not being applied
	- \#40 Ensure non-string data types (numeric, date, boolean) are respected when processing cells
- Enhancements:
	- \#17 Add `setActiveSheetNumber()`
	- \#18 Add `formatRows()`
	- \#19 Support reading sheets by name
	- \#20 Add `deleteRows()`
	- \#21 Add `deleteColumn()` and `deleteColumns()`
	- \#22 Add `shiftColumns()`
	- \#23 Add `getCellValue()` and `setCellValue()`
	- \#24 Add `formatColumn()`, `formatColumns()` and `formatCellRange()`
	- \#23 Add `isBinaryFormat()` and `isXmlFormat()`
	- \#28 Add `mergeCells()`
	- \#29 Add `addFreezePane()` and `addSplitPane()`
	- \#30 Add `addInfo()` and `info()`
	- \#31 Add `setCellFormula()` and `getCellFormula()`
	- \#34 Add `setColumnWidth()`
	- \#35 Add `setRowHeight()`
	- \#36 Add `setHeader()` and `setFooter()`
	- \#33 Add `setCellComment()` and `getCellComment()`
	- \#32 Add `addImage()`
	- \#37 Add `autoSizeColumn()`
	- \#41 Add option to auto size columns when using addColumn, addRow and addRow
	- \#39 Add `renameSheet()`
	- \#38 Add `clearCell()` and `clearCellRange()`

## 0.2.0 - 18 February 2015
- Breaking changes
 - `read()` method `sheet` argument should now be `sheetNumber` (for consistency)
 - When specifying 1-based sheet numbers as arguments, always use `sheetNumber` (not `sheet` or `sheetIndex`).
 - When specifying sheet names as arguments, use `sheetName`, not `sheet`.
- Enhancements
 - \#13 Add support for `createSheet()`
 - \#14 Add support for `removeSheet()`
 - \#15 Add `deleteSheet()` which can delete a sheet by name or number

## 0.1.0 - 17 February 2015
- Bug fixes:
 - Treat null rows/cells as blank not null
 - \#5 `new()` method ignores xmlFormat argument
 - \#6 `ShiftRows` offset argument misspel
 - \#7 `ShiftRows` calls require workbook as argumen
 - \#8 `AddRow` insert argument not working
 - \#10 Cannot read XLSX files
 - \#11 `Read` method errors if no format specified. Should return workbook object
- Enhancements:
 - \#2 Testbox BDD style test suite
 - \#3 Upgrade POI to 3.11
 - \#3 Option to include blank rows when reading into a query
 - \#4 Simplify dependencies by including tools and formatting as mixins
 - \#9 `writeFileFromQuery()`: detect if xml from file extension
 - \#12 Change ACF `excludeHeaderRow` default=false to `includeHeaderRow`, default=false 

## 0.0.5 - 13 February 2015
- `read` method
	- changed radically to work under Lucee. Some attributes/functionality disabled for now, but can return a query or workbook object.
	- changed `excludeHeaderRow` default from false to true
- Added `flushPoiLoader` utility method

## 0.0.4 - 12 February 2015
- Added
 - `write` method matching `SpreadSheetWrite()`
 - `writeFileFromQuery` custom method

## 0.0.3 - 25 January 2015
- Workbook creation separated from instantiation. Create a workbook using `new()` and then pass it to other functions. Same as ACF functions.
- Use JavaLoader to load newer POI jars to allow support for `read()`
- Added methods
 - `new`
 - `read` (matches `cfspreadsheet action="read"`)
 - `setActiveSheet`

## 0.0.2 - 19 January 2015
- Added custom method: `downloadFileFromQuery`

## 0.0.1 - 18 January 2015
- Initial release with support for the following standard CFML functions only:
	- `addColumn`
	- `addRow`
	- `addRows`
	- `deleteRow`
	- `formatCell`
	- `formatRow`
	- `shiftRows`
	- `readBinary`
- Custom method: `binaryFromQuery`