# Spreadsheet CFML

Standalone library for working with spreadsheets and CSV in CFML ([Lucee](http://lucee.org/) and Adobe ColdFusion), supporting all of ColdFusion's native spreadsheet functionality and much more besides.

## Minimum Requirements

- Java 8 or higher
- Lucee 5.x or higher
- Adobe ColdFusion 2018 or higher

## Installation

Note that this is not an extension or package, so **does not need to be installed**. To use it, simply copy the files/folders to a location where `Spreadsheet.cfc` can be called by your application code.

The following are the essential files/folders you will need depending on which CFML engine you are using:

### Lucee
```
helpers/
objects/
lib-osgi.jar
osgiLoader.cfc
Spreadsheet.cfc
```
### Adobe ColdFusion
```
helpers/
javaLoader/
lib/
objects/
Spreadsheet.cfc
```

## Usage

The following example assumes the file containing the script is in the same directory as the folder containing the spreadsheet library files, i.e.:

- root/
  - spreadsheetLibrary/
    - Spreadsheet.cfc
    - etc.
  - script.cfm
 
```
<cfscript>
spreadsheet = New spreadsheetLibrary.Spreadsheet();
data = QueryNew( "First,Last", "VarChar, VarChar", [ [ "Susi", "Sorglos" ], [ "Frumpo", "McNugget" ] ] );
workbook = spreadsheet.new();
spreadsheet.addRows( workbook, data );
spreadsheet.write( workbook, "c:/temp/data.xls" );
</cfscript>
```
### init()

When instantiating the library, the `init()` method **must** be called. This will happen automatically if you use the `New` keyword:
```
spreadsheet = New spreadsheetLibrary.Spreadsheet();
```
If using `CreateObject()` then you must call `init()` explicitly:
```
spreadsheet = CreateObject( "component", "spreadsheetLibrary.Spreadsheet" ).init();
```
### Using a mapping

You may wish to place the spreadsheet library files in a central location with an application mapping, and instantiate the component using its dot path (e.g. `New myLibrary.spreadsheet.Spreadsheet();`).

[How to create mappings (StackOverflow)](http://stackoverflow.com/questions/12073714/components-mapping-in-railo).

[Full function reference](https://github.com/cfsimplicity/spreadsheet-cfml/wiki)

## Supported ColdFusion functions

* [addAutofilter](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/addAutofilter)
* [addColumn](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/addColumn)
* [addFreezePane](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/addFreezePane)
* [addImage](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/addImage)
* [addInfo](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/addInfo)
* [addPageBreaks](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/addPageBreaks)
* [addRow](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/addRow)
* [addRows](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/addRows)
* [addSplitPane](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/addSplitPane)
* autosize, implemented as [autoSizeColumn](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/autoSizeColumn)
* [createSheet](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/createSheet)
* [deleteColumn](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/deleteColumn)
* [deleteColumns](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/deleteColumns)
* [deleteRow](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/deleteRow)
* [deleteRows](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/deleteRows)
* [formatCell](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/formatCell)
* [formatCellRange](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/formatCellRange)
* [formatColumn](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/formatColumn)
* [formatColumns](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/formatColumns)
* [formatRow](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/formatRow)
* [formatRows](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/formatRows)
* [getCellComment](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/getCellComment)
* [getCellFormula](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/getCellFormula)
* [getCellValue](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/getCellValue)
* [getColumnCount](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/getColumnCount)
* [info](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/info)
* [isSpreadsheetFile](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/isSpreadsheetFile)
* [isSpreadsheetObject](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/isSpreadsheetObject)
* [mergeCells](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/mergeCells)
* [new](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/new)
* [read](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/read)
* [readBinary](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/readBinary)
* [removeSheet](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/removeSheet)
* [setActiveSheet](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setActiveSheet)
* [setActiveSheetNumber](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setActiveSheetNumber)
* [setCellComment](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setCellComment)
* [setCellFormula](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setCellFormula)
* [setCellValue](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setCellValue)
* [setColumnWidth](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setColumnWidth)
* [setFooter](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setFooter)
* [setHeader](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setHeader)
* [setRowHeight](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setRowHeight)
* [shiftColumns](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/shiftColumns)
* [shiftRows](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/shiftRows)
* [write](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/write)

### Extra functions not available in ColdFusion

* [addConditionalFormatting](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/addConditionalFormatting)
* [addDataValidation](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/addDataValidation)
* [addPrintGridlines](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/addPrintGridlines)
* [binaryFromQuery](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/binaryFromQuery)
* [cleanUpStreamingXml](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/cleanUpStreamingXml)
* [clearCell](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/clearCell)
* [clearCellRange](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/clearCellRange)
* [createCellStyle](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/createCellStyle)
* [csvToQuery](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/csvToQuery)
* [download](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/download)
* [downloadCsvFromFile](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/downloadCsvFromFile)
* [downloadFileFromQuery](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/downloadFileFromQuery)
* [getCellAddress](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/getCellAddress)
* [getCellComments](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/getCellComments)
* [getCellFormat](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/getCellFormat)
* [getCellHyperLink](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/getCellHyperLink)
* [getCellType](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/getCellType)
* [getColumnWidth](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/getColumnWidth)
* [getColumnWidthInPixels](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/getColumnWidthInPixels)
* [getLastRowNumber](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/getLastRowNumber)
* [getPOIVersion](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/getPOIVersion)
* [getPresetColorNames](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/getPresetColorNames)
* [getRowCount](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/getRowCount)
* [hideColumn](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/hideColumn)
* [hideRow](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/hideRow)
* [isBinaryFormat](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/isBinaryFormat)
* [isColumnHidden](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/isColumnHidden)
* [isRowHidden](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/isRowHidden)
* [isStreamingXmlFormat](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/isStreamingXmlFormat)
* [isXmlFormat](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/isXmlFormat)
* [newStreamingXlsx](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/newStreamingXlsx)
* [newXls](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/newXls)
* [newXlsx](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/newXlsx)
* [queryToCsv](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/queryToCsv)
* [readCsv](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/readCsv)
* [readLargeFile](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/readLargeFile)
* [removePrintGridlines](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/removePrintGridlines)
* [renameSheet](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/renameSheet)
* [removeSheetNumber](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/removeSheetNumber)
* [setActiveCell](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setActiveCell)
* [setCellHyperLink](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setCellHyperLink)
* [setCellRangeValue](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setCellRangeValue)
* [setDateFormats](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setDateFormats)
* [setDefaultWorkbookFormat](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setDefaultWorkbookFormat)
* [setFitToPage](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setFitToPage)
* [setFooterImage](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setFooterImage)
* [setHeaderImage](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setHeaderImage)
* [setReadOnly](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setReadOnly)
* [setRecalculateFormulasOnNextOpen](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setRecalculateFormulasOnNextOpen)
* [setRepeatingColumns](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setRepeatingColumns)
* [setRepeatingRows](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setRepeatingRows)
* [setReturnCachedFormulaValues](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setReturnCachedFormulaValues)
* [setSheetTopMargin](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setSheetTopMargin)
* [setSheetBottomMargin](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setSheetBottomMargin)
* [setSheetLeftMargin](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setSheetLeftMargin)
* [setSheetRightMargin](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setSheetRightMargin)
* [setSheetHeaderMargin](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setSheetHeaderMargin)
* [setSheetFooterMargin](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setSheetFooterMargin)
* [setSheetPrintOrientation](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/setSheetPrintOrientation)
* [sheetInfo](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/sheetInfo)
* [showColumn](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/showColumn)
* [showRow](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/showRow)
* [workbookFromCsv](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/workbookFromCsv)
* [workbookFromQuery](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/workbookFromQuery)
* [writeFileFromQuery](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/writeFileFromQuery)
* [writeCsv](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/writeCsv)
* [writeToCsv](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/writeToCsv)

### Enhanced Read() method

In Adobe ColdFusion, the `SpreadsheetRead()` script function is limited to just returning a spreadsheet object, whereas the `<cfspreadsheet action="read">` tag has a range of options for reading and returning data from a spreadsheet file. 

The `read()` method in this library allows you to read a spreadsheet file into a query and return that instead of a spreadsheet object. It includes all of the options available in `<cfspreadsheet action="read">`.

```
<cfscript>
myQuery = spreadsheet.read( src=mypath, format="query" );
</cfscript>
```

The `read()` method also features the following additional options not available in ColdFusion or the Spreadsheet Extension:

* `fillMergedCellsWithVisibleValue`
* `includeHiddenColumns`
* `includeRichTextFormatting`
* `password` to open encrypted spreadsheets
* `csvDelimiter`
* `queryColumnTypes`

[Full documentation of read()](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/read)

### Chainable syntax

From version 3, multiple calls can be chained together, simplifying your code into a more expressive syntax.
```
spreadsheet.newChainable( "xlsx" )
 .addRows( data )
 .formatRow( { bold: true }, 1 )
 .write( filepath );
```
[More on chaining calls](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/chaining)

### Date formats

The following international date masks are used by default to read and write cell values formatted as dates:

* DATE = `yyyy-mm-dd`
* TIME = `hh:mm:ss`
* TIMESTAMP = `yyyy-mm-dd hh:mm:ss`

An additional mask is used to output datetime values from the `read()` method into HTML or CSV formats:

* DATETIME = `yyyy-mm-dd HH:nn:ss`

NB: _Do not confuse `DATETIME` and `TIMESTAMP`._ In general you should override the `TIMESTAMP` mask.

Each of these can be overridden by passing in a struct including the value(s) to be overridden when instantiating the Spreadsheet component. For example:
```
<cfscript>
spreadsheet = New spreadsheetLibrary.Spreadsheet( dateFormats={ DATE: "mm/dd/yyyy" } );
</cfscript>
```
Or by using the `setDateFormats()` method on an existing instance.
```
<cfscript>
spreadsheet = New spreadsheetLibrary.Spreadsheet();
spreadsheet.setDateFormats( { DATE: "mm/dd/yyyy" } );
</cfscript>
```
While the above will set the library defaults, you can format cells with specific masks using the `dataFormat` attribute which can be passed to [formatCell](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/formatCell) and the other [formatting](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/Formatting-options) methods, as part of the `format` argument:
```
// display datetime value with millisecond precision
spreadsheet.formatColumn( workbook , { dataformat: "yyyy-mm-dd hh:mm:ss.000" }, 1 );
```
### JavaLoader

From version 2.14.0, Lucee loads the POI and other required java libraries using OSGi. This is not yet supported with Adobe ColdFusion which by default uses an included version of Mark Mandel's [JavaLoader](https://github.com/markmandel/JavaLoader). 

For more details and options see: [Loading the POI java libraries](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/Loading-the-POI-java-libraries)

## CommandBox Installation

You can also download this library through CommandBox/Forgebox.
```
box install spreadsheet-cfml
```
It will download the files into a modules directory and can be used just the same as downloading the files manually.

If using ColdBox you can use either of the WireBox bindings like so:
```
spreadsheet = wirebox.getInstance( "Spreadsheet@spreadsheet-cfml" );
spreadsheet = wirebox.getInstance( "Spreadsheet CFML" );
```

## Test Suite

The automated tests require [TestBox 5.0](https://github.com/Ortus-Solutions/TestBox) or later. You will need to create an application mapping for `/testbox`

## Credits

The code was originally adapted from the work of [TeamCfAdvance](https://github.com/teamcfadvance/). Ben Nadel's [POI Utility](https://github.com/bennadel/POIUtility.cfc) was also used as a basis for parts of the `read` functionality. Header/Footer image functionality is based on code by [Axel Richter](https://stackoverflow.com/users/3915431/axel-richter).

[JavaLoader](https://github.com/markmandel/JavaLoader) is by Mark Mandel.

## Legal

### The MIT License (MIT)

Copyright (c) 2015-2024 Julian Halliwell

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
