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