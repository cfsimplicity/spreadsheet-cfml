#Spreadsheet library for Railo 4.2

Adapted from the https://github.com/teamcfadvance/cfspreadsheet-railo extension, this is a standalone library for reading, creating and formatting spreadsheets in Railo 4.2 which does not require installation into each web context.

##Rationale

I was dissatisfied with the official Railo spreadsheet extension for two main reasons:

1. It was designed for an older version of Railo and (at the time of writing) installation as an extension fails in version 4.2
2. It can be installed manually, but this needs doing in each web context, followed by a server restart

##Benefits over the official extension

- No installation/restart required, either at the server or individual web context level.
- `read()` method offers features of `<cfspreadsheet action="read">` tag in script rather than the more limited options with `SpreadsheetNew()`.
- Invoking the library doesn't create a workbook instance (a.k.a *Spreadsheet Object*), meaning:
 - a blank workbook isn't created unnecessarily when reading an existing spreadsheet
 - the library can be stored as a singleton in application scope
 - the functions work more like those in ACF: you pass in an existing workbook explicitly as the first argument.
- Offers additional convenience methods, e.g. `downloadFileFromQuery()`.
- Written entirely in Railo 4.2 script.

##Downsides

- Not all spreadsheet functions implemented
- Existing code needs adapting to invoke the library. Existing CFML spreadsheet functions and the `<cfspreadsheet>` tag won't work with it.

##Currently supported standard functions

- `addColumn`
- `addRow`
- `addRows`
- `deleteRow`
- `formatCell`
- `formatRow`
- `new`
- `read`
- `readBinary`
- `setActiveSheet`
- `shiftRows`

##Usage

```
spreadsheet	=	New spreadsheet();
data = QueryNew( "First,Last","VarChar,VarChar",[ [ "Susi","Sorglos" ],[ "Julian","Halliwell" ] ] );
workbook = spreadsheet.new();
spreadsheet.addRows( workbook,data );
```

###Enhanced Read() method

In Adobe ColdFusion, the `SpreadsheetRead()` script function is limited to just returning a spreadsheet object, whereas the `<cfspreadsheet action="read">` tag has a range of options for reading and returning data from a spreadsheet file. The `read()` method in this library can take the `cfspreadsheet` attributes as arguments, with the exception of the `query` attribute which is unnecessary in script. To return a query simply specify "query" in the `format` argument:

```
myQuery = spreadsheet.read( src=mypath,format="query" );
```

###Convenience methods

####downloadFileFromQuery()

Provides a quick way of downloading a spreadsheet to the browser by passing a query and a filename. The query column names are included by default as a bold header row.

```
spreadsheet	=	New spreadsheet();
data = QueryNew( "First,Last","VarChar,VarChar",[ [ "Susi","Sorglos" ],[ "Julian","Halliwell" ] ] );
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

####binaryFromQuery()
Similar to `downloadFileFromQuery`, but without downloading the file.

```
data = QueryNew( "First,Last","VarChar,VarChar",[ [ "Susi","Sorglos" ],[ "Julian","Halliwell" ] ] );
spreadsheet	=	New spreadsheet();
binary = spreadsheet.binaryFromQuery( data );
```

##Credits

The code is very largely based on the work of [TeamCfAdvance](https://github.com/teamcfadvance/), to whom credit and thanks are due.

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