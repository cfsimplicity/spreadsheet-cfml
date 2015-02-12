## 0.0.4 - 12 February

- Adds
 - `write` method matching `SpreadSheetWrite()`
 - `writeFileFromQuery` custom method


## 0.0.3 - 25 January 2015
- Workbook creation separated from instantiation. Create a workbook using `new()` and then pass it to other functions. Same as ACF functions.
- Use JavaLoader to load newer POI jars to allow support for `read()`
- Adds methods
 - `new`
 - `read` (matches `cfspreadsheet action="read"`)
 - `setActiveSheet`

## 0.0.2 - 19 January 2015
Adds custom method: `downloadFileFromQuery`

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