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