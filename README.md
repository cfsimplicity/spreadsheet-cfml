#Spreadsheet library for Railo 4.x

Adapted from the https://github.com/teamcfadvance/cfspreadsheet-railo extension, this is a standalone library for creating and working with spreadsheets in Railo 4.x which does not require installation into each web context.

##Rationale

I was dissatisfied with the official Railo spreadsheet extension for two main reasons:

1. It was designed for an older version of Railo and (at the time of writing) installation as an extension fails in version 4.x
2. It can be installed manually, but this needs doing in each web context, followed by a server restart

##Benefits over the official extension

- No installation required, either at the server or individual web context level.
- No additional java classes need installing/loading: it uses jars already loaded by the core Railo 4.x engine.
- Offers additional convenience methods, e.g. `binaryFromQuery()`.
- Written entirely in Railo 4.x script.

##Downsides

- Currently only a limited sub-set of functions. More will be implemented in due course.
- Existing code needs adapting to invoke the library. Existing function calls won't work with it unlike the extension.
- No tag support.

##Usage

```
data = QueryNew( "First,Last","VarChar,VarChar",[ [ "Susi","Sorglos" ],[ "Julian","Halliwell" ] ] );
spreadsheet	=	New Spreadsheet();
spreadsheet.addRows( data );
binary = spreadsheet.readBinary();
header name="Content-Disposition" value="attachment; filename=#Chr( 34 )#report.xls#Chr( 34 )#";
content type="application/msexcel" variable="#binary#" reset="true";
```

###`binaryFromQuery()`.

Provices a quick way of transforming a query into a downloadable spreadsheet with the column names as a header row

```
data = QueryNew( "First,Last","VarChar,VarChar",[ [ "Susi","Sorglos" ],[ "Julian","Halliwell" ] ] );
spreadsheet	=	New Spreadsheet();
binary = spreadsheet.binaryFromQuery( data );
header name="Content-Disposition" value="attachment; filename=#Chr( 34 )#report.xls#Chr( 34 )#";
content type="application/msexcel" variable="#binary#" reset="true";
```

##Credits

The code is very largely based on the work of [TeamCfAdvance](https://github.com/teamcfadvance/), to whom credit and thanks are due.