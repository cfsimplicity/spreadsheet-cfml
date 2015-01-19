#Spreadsheet library for Railo 4.x

Adapted from the https://github.com/teamcfadvance/cfspreadsheet-railo extension, this is a standalone library for creating and formatting spreadsheets in Railo 4.x which does not require installation into each web context.

##Rationale

I was dissatisfied with the official Railo spreadsheet extension for two main reasons:

1. It was designed for an older version of Railo and (at the time of writing) installation as an extension fails in version 4.x
2. It can be installed manually, but this needs doing in each web context, followed by a server restart

##Benefits over the official extension

- No installation required, either at the server or individual web context level.
- No additional java classes need installing/loading: it uses jars already loaded by the core Railo 4.x engine.
- Offers additional convenience methods, e.g. `downloadFileFromQuery()`.
- Written entirely in Railo 4.x script.

##Downsides

- Limited sub-set of functions for generating spreadsheets only.
- Existing code needs adapting to invoke the library. Existing CFML spreadsheet functions and the `<cfspreadsheet>` tag won't work with it.

##Currently supported standard functions

- `addColumn`
- `addRow`
- `addRows`
- `deleteRow`
- `formatCell`
- `formatRow`
- `shiftRows`
- `readBinary`

##Usage

```
data = QueryNew( "First,Last","VarChar,VarChar",[ [ "Susi","Sorglos" ],[ "Julian","Halliwell" ] ] );
spreadsheet	=	New Spreadsheet();
spreadsheet.addRows( data );
```

To specify the sheet name, include it when instantiating the spreadsheet:

```
spreadsheet	=	New Spreadsheet( "CustomSheetName" );
```

###Convenience methods

####downloadFileFromQuery

Provides a quick way of downloading a spreadsheet to the browser by passing a query and a filename. The query column names are included by default as a bold header row.

```
data = QueryNew( "First,Last","VarChar,VarChar",[ [ "Susi","Sorglos" ],[ "Julian","Halliwell" ] ] );
spreadsheet	=	New Spreadsheet();
filename = "report";
spreadsheet.downloadFileFromQuery( data,filename );
```

If you don't want the header row:

```
spreadsheet.downloadFileFromQuery( data,filename,addHeaderRow=false );
```

If you want the header row, but not bold:

```
spreadsheet.downloadFileFromQuery( data,filename,boldHeaderRow=false );
```

####binaryFromQuery
Similar to `downloadFileFromQuery`, but without downloading the file.

```
data = QueryNew( "First,Last","VarChar,VarChar",[ [ "Susi","Sorglos" ],[ "Julian","Halliwell" ] ] );
spreadsheet	=	New Spreadsheet();
binary = spreadsheet.binaryFromQuery( data );
```

##Credits

The code is very largely based on the work of [TeamCfAdvance](https://github.com/teamcfadvance/), to whom credit and thanks are due.

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