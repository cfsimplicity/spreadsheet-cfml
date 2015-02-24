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
Added custom method: `downloadFileFromQuery`

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