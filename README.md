#Spreadsheet library for Lucee

Adapted from the https://github.com/teamcfadvance/cfspreadsheet-railo extension, this is a standalone library for reading, creating and formatting spreadsheets in [Lucee Server](http://lucee.org/) which does not require installation into each web context.

##Rationale

I was dissatisfied with the official Railo (now Lucee) spreadsheet extension for two main reasons:

1. It was designed for an older version of Railo (Lucee) and (at the time of writing) installation as an extension fails in the current version.
2. It can be installed manually, but this is cumbersome and needs doing in each web context followed by a server restart.

##Benefits over the official extension

- No installation/restart required, either at the server or individual web context level.
- Invoking the library doesn't create a workbook instance (a.k.a *Spreadsheet Object*), meaning:
 - a blank workbook isn't created unnecessarily when reading an existing spreadsheet
 - the library can be stored as a singleton in application scope
 - the functions work more like those in ACF: you pass in an existing workbook explicitly as the first argument.
- `read()` method offers all the features of the `<cfspreadsheet action="read">` tag in script in addition to the basic options of `SpreadsheetRead()`.
- Offers additional convenience methods, e.g. `downloadFileFromQuery()`.
- Uses a newer version of the java POI library (3.12).
- Fixes various outstanding bugs/omissions.
- No dependency on Railo/Lucee within the included jar files.
- Written entirely in Lucee script.

##Downsides

- Existing code needs adapting to invoke the library. Existing CFML spreadsheet functions and the `<cfspreadsheet>` tag won't work with it.
- Populating a spreadsheet from CSV is not supported.

##Usage

Note that this is not a Lucee extension, so **does not need to be installed**. To use it, simply copy the files/folders to a location where `Spreadsheet.cfc` can be called by your application code.

The following example assumes the file containing the script is in the same location as the spreadsheet.cfc.
```
spreadsheet	=	New spreadsheet();
data = QueryNew( "First,Last","VarChar,VarChar",[ [ "Susi","Sorglos" ],[ "Julian","Halliwell" ] ] );
workbook = spreadsheet.new();
spreadsheet.addRows( workbook,data );
```
You will probably want to place the spreadsheet library files in a central location with an application mapping, and instantiate the component using its dot path (e.g. `New myLibrary.spreadsheet.spreadsheet();`).

[How to create mappings (StackOverflow)](http://stackoverflow.com/questions/12073714/components-mapping-in-railo).

[Full function reference](https://github.com/cfsimplicity/lucee-spreadsheet/wiki)

##Supported ColdFusion 9 functions

* [addColumn](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/addColumn)
* [addFreezePane](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/addFreezePane)
* [addImage](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/addImage)
* [addInfo](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/addInfo)
* [addRow](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/addRow)
* [addRows](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/addRows)
* [addSplitPane](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/addSplitPane)
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
* [info](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/info)
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

###Extra functions not available in ColdFusion

* [autoSizeColumn](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/autoSizeColumn)
* [clearCell](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/clearCell)
* [clearCellRange](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/clearCellRange)
* [hideColumn](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/hideColumn)
* [isBinaryFormat](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/isBinaryFormat)
* [isXmlFormat](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/isXmlFormat)
* [renameSheet](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/renameSheet)
* [removeSheetNumber](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/removeSheetNumber)
* [setCellRangeValue](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/setCellRangeValue)
* [showColumn](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/showColumn)

###Additional Convenience methods

* [binaryFromQuery](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/binaryFromQuery)
* [downloadCsvFromFile](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/downloadCsvFromFile)
* [downloadFileFromQuery](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/downloadFileFromQuery)
* [workbookFromQuery](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/workbookFromQuery)
* [writeFileFromQuery](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/writeFileFromQuery)

###Enhanced Read() method

In Adobe ColdFusion, the `SpreadsheetRead()` script function is limited to just returning a spreadsheet object, whereas the `<cfspreadsheet action="read">` tag has a range of options for reading and returning data from a spreadsheet file. 

The `read()` method in this library allows you to read a spreadsheet file into a query and return that instead of a spreadsheet object. It includes all of the options available in `<cfspreadsheet action="read">`.

```
myQuery = spreadsheet.read( src=mypath,format="query" );
```

The `read()` method also features the following additional options not available in ColdFusion or the Spreadsheet Extension:

* `fillMergedCellsWithVisibleValue`
* `includeHiddenColumns`
* `includeRichTextFormatting`

[Full documentation of read()](https://github.com/cfsimplicity/lucee-spreadsheet/wiki/read)

###Date formats

The following international date masks are used by default to read and write cell values formatted as dates:

* DATE=`yyyy-mm-dd`
* DATETIME=`yyyy-mm-dd HH:nn:ss`
* TIME=`hh:mm:ss`
* TIMESTAMP=`yyyy-mm-dd hh:mm:ss`

Each of these can be overridden by passing in a struct including the value(s) to be overridden when instantiating the Spreadsheet component. For example:

```
spreadsheet = New root.spreadsheet( dateFormats={ DATE="mm/dd/yyyy" } );
```

##Test Suite
The automated tests require [TestBox 2.1](https://github.com/Ortus-Solutions/TestBox). You will need to create an application mapping for `/testbox`

##Credits

The code is very largely based on the work of [TeamCfAdvance](https://github.com/teamcfadvance/), to whom credit and thanks are due. Ben Nadel's [POI Utility](https://github.com/bennadel/POIUtility.cfc) was also used as a basis for parts of the `read` functionality.

[JavaLoader](https://github.com/markmandel/JavaLoader) is by Mark Mandel.

##Legal

###The MIT License (MIT)

Copyright (c) 2015 Julian Halliwell

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