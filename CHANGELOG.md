## 0.1.0 - 17 February 2015

- Bug fixes:
 - Treat null rows/cells as blank not null
 - \#5 new() method ignores xmlFormat argument
 - \#6 ShiftRows offset argument misspel
 - \#7 ShiftRows calls require workbook as argumen
 - \#8 AddRow insert argument not working
 - \#10 Cannot read XLSX files
 - \#11 Read method errors if no format specified. Should return workbook object
- Enhancements:
 - \#2 Testbox BDD style test suite
 - \#3 Upgrade POI to 3.11
 - \#3 Option to include blank rows when reading into a query
 - \#4 Simplify dependencies by including tools and formatting as mixins
 - \#9 writeFileFromQuery(): detect if xml from file extension
 - \#12 Change ACF "excludeHeaderRow" default=false to "includeHeaderRow", default=false 

## 0.0.5 - 13 February 2015

- `read` method
	- changed radically to work under Lucee. Some attributes/functionality disabled for now, but can return a query or workbook object.
	- changed excludeHeaderRow default from false to true
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