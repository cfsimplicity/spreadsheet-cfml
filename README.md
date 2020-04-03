# Spreadsheet library for Lucee (and Adobe ColdFusion)

Originally adapted from the https://github.com/teamcfadvance/cfspreadsheet-railo extension, this is a standalone library for reading, creating and formatting spreadsheets in [Lucee Server](http://lucee.org/) which does not require installation as an extension. It can also be used with Adobe ColdFusion.

## Rationale

Unlike Adobe ColdFusion, Lucee doesn't support spreadsheet functionality out of the box. An [extension exists for Lucee](https://github.com/Leftbower/cfspreadsheet-lucee-5), but I decided to create a standalone library which doesn't depend on customisation of the engine.

## Library vs Extension

### Benefits

- No installation required, either at the server or individual web context level.
- `read()` method offers all the features of the `<cfspreadsheet action="read">` tag in script in addition to the basic options of `SpreadsheetRead()`.
- Offers a number of additional functions and options (see below)
- Fixes various outstanding bugs/omissions.
- No dependency on Lucee within the included jar files.
- Invoking the library doesn't create a workbook instance (a.k.a. *Spreadsheet Object*), meaning:
  - a blank workbook isn't created unnecessarily when reading an existing spreadsheet
  - the library can be stored as a singleton in application scope
- Also works with Adobe ColdFusion (see below)
- Written entirely in CFML script.

### Downsides

- Existing code needs adapting to invoke the library. Existing CFML spreadsheet functions and the `<cfspreadsheet>` tag won't work with it.

### Adobe ColdFusion

Although primarily intended for Lucee, the library can be run under ColdFusion 2016 or higher. This may be useful where you want to your codebase to be cross-compatible between the two engines.

Please note though that _writing password-protected files only works with Lucee_.

## Minimum Requirements for version 2.x

- Java 8 or higher
- Lucee 5.x or higher
- Adobe ColdFusion 2016 or higher

If you are running Java 6 and 7, Lucee 4.5 or ACF11, please use [version 1.x.](https://github.com/cfsimplicity/lucee-spreadsheet/tags)

## Usage

Note that this is not a Lucee extension, so **does not need to be installed**. To use it, simply copy the files/folders to a location where `Spreadsheet.cfc` can be called by your application code.

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
</cfscript>
```
You will probably want to place the spreadsheet library files in a central location with an application mapping, and instantiate the component using its dot path (e.g. `New myLibrary.spreadsheet.Spreadsheet();`).

[How to create mappings (StackOverflow)](http://stackoverflow.com/questions/12073714/components-mapping-in-railo).

[Full function reference](https://github.com/cfsimplicity/lucee-spreadsheet/wiki)

## Supported ColdFusion functions

* [addAutofilter](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/addAutofilter)
* [addColumn](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/addColumn)
* [addFreezePane](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/addFreezePane)
* [addImage](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/addImage)
* [addInfo](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/addInfo)
* [addPageBreaks](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/addPageBreaks)
* [addRow](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/addRow)
* [addRows](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/addRows)
* [addSplitPane](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/addSplitPane)
* autosize, implemented as [autoSizeColumn](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/autoSizeColumn)
* [createSheet](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/createSheet)
* [deleteColumn](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/deleteColumn)
* [deleteColumns](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/deleteColumns)
* [deleteRow](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/deleteRow)
* [deleteRows](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/deleteRows)
* [formatCell](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/formatCell)
* [formatCellRange](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/formatCellRange)
* [formatColumn](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/formatColumn)
* [formatColumns](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/formatColumns)
* [formatRow](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/formatRow)
* [formatRows](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/formatRows)
* [getCellComment](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/getCellComment)
* [getCellFormula](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/getCellFormula)
* [getCellValue](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/getCellValue)
* [getColumnCount](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/getColumnCount)
* [info](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/info)
* [isSpreadsheetFile](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/isSpreadsheetFile)
* [isSpreadsheetObject](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/isSpreadsheetObject)
* [mergeCells](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/mergeCells)
* [new](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/new)
* [read](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/read)
* [readBinary](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/readBinary)
* [removeSheet](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/removeSheet)
* [setActiveSheet](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setActiveSheet)
* [setActiveSheetNumber](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setActiveSheetNumber)
* [setCellComment](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setCellComment)
* [setCellFormula](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setCellFormula)
* [setCellValue](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setCellValue)
* [setColumnWidth](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setColumnWidth)
* [setFooter](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setFooter)
* [setHeader](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setHeader)
* [setRowHeight](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setRowHeight)
* [shiftColumns](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/shiftColumns)
* [shiftRows](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/shiftRows)
* [write](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/write)

### Extra functions not available in ColdFusion

* [addPrintGridlines](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/addPrintGridlines)
* [cleanUpStreamingXml](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/cleanUpStreamingXml)
* [clearCell](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/clearCell)
* [clearCellRange](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/clearCellRange)
* [getCellFormat](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/getCellFormat)
* [getCellType](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/getCellType)
* [getColumnWidth](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/getColumnWidth)
* [getColumnWidthInPixels](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/getColumnWidthInPixels)
* [getPresetColorNames](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/getPresetColorNames)
* [getRowCount](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/getRowCount)
* [hideColumn](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/hideColumn)
* [hideRow](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/hideRow)
* [isBinaryFormat](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/isBinaryFormat)
* [isColumnHidden](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/isColumnHidden)
* [isRowHidden](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/isRowHidden)
* [isStreamingXmlFormat](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/isStreamingXmlFormat)
* [isXmlFormat](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/isXmlFormat)
* [removePrintGridlines](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/removePrintGridlines)
* [renameSheet](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/renameSheet)
* [removeSheetNumber](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/removeSheetNumber)
* [setCellRangeValue](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setCellRangeValue)
* [setFitToPage](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setFitToPage)
* [setReadOnly](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setReadOnly)
* [setRepeatingColumns](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setRepeatingColumns)
* [setRepeatingRows](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setRepeatingRows)
* [setSheetTopMargin](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setSheetTopMargin)
* [setSheetBottomMargin](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setSheetBottomMargin)
* [setSheetLeftMargin](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setSheetLeftMargin)
* [setSheetRightMargin](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setSheetRightMargin)
* [setSheetHeaderMargin](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setSheetHeaderMargin)
* [setSheetFooterMargin](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setSheetFooterMargin)
* [setSheetPrintOrientation](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setSheetPrintOrientation)
* [showColumn](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/showColumn)
* [showRow](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/showRow)

### Additional Convenience methods

* [binaryFromQuery](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/binaryFromQuery)
* [csvToQuery](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/csvToQuery)
* [download](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/download)
* [downloadCsvFromFile](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/downloadCsvFromFile)
* [downloadFileFromQuery](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/downloadFileFromQuery)
* [newStreamingXlsx](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/newStreamingXlsx)
* [newXls](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/newXls)
* [newXlsx](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/newXlsx)
* [workbookFromCsv](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/workbookFromCsv)
* [workbookFromQuery](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/workbookFromQuery)
* [writeFileFromQuery](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/writeFileFromQuery)

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

[Full documentation of read()](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/read)

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
spreadsheet = New spreadsheetLibrary.spreadsheet( dateFormats={ DATE: "mm/dd/yyyy" } );
</cfscript>
```
While the above will set the library defaults, you can format cells with specific masks using the `dataFormat` attribute which can be passed to [formatCell](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/formatCell) and the other [formatting](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/Formatting-options) methods, as part of the `format` argument:
```
// display datetime value with millisecond precision
spreadsheet.formatColumn( workbook , { dataformat: "yyyy-mm-dd hh:mm:ss.000" }, 1 );
```

### JavaLoader

A bundled version of Mark Mandel's [JavaLoader](https://github.com/markmandel/JavaLoader) will be used by default to load the POI and other required java libraries.

For more details and options see: [Loading the POI java libraries](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/Loading-the-POI-java-libraries)

## CommandBox Installation

You can also download this library through CommandBox.
```
box install cfsimplicity/lucee-spreadsheet
```
It will download the files into a modules directory and can be used just the same as downloading the files manually.

If using ColdBox you can use either of the WireBox bindings like so:
```
spreadsheet = wirebox.getInstance("Spreadsheet@lucee-spreadsheet");
spreadsheet = wirebox.getInstance("LuceeSpreadsheet");
```

## Test Suite
The automated tests require [TestBox 2.1](https://github.com/Ortus-Solutions/TestBox) or later. You will need to create an application mapping for `/testbox`

## Credits

The code was originally adapted from the work of [TeamCfAdvance](https://github.com/teamcfadvance/). Ben Nadel's [POI Utility](https://github.com/bennadel/POIUtility.cfc) was also used as a basis for parts of the `read` functionality.

[JavaLoader](https://github.com/markmandel/JavaLoader) is by Mark Mandel.

## Legal

### The MIT License (MIT)

Copyright (c) 2015-2020 Julian Halliwell

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