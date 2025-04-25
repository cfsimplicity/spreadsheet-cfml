# Spreadsheet CFML

Standalone library for working with spreadsheets and CSV in CFML ([Lucee](http://lucee.org/) and Adobe ColdFusion), [supporting all of ColdFusion's native spreadsheet functionality](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/ColdFusion-spreadsheet-functionality-support) and [much more](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/Functions-not-available-in-ColdFusion) besides.

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
  - spreadsheetCFML/
    - Spreadsheet.cfc
    - etc.
  - script.cfm
```
<cfscript>
spreadsheet = New spreadsheetCFML.Spreadsheet();
data = QueryNew( "First,Last", "VarChar, VarChar", [ [ "Susi", "Sorglos" ], [ "Frumpo", "McNugget" ] ] );
workbook = spreadsheet.new();
spreadsheet.addRows( workbook, data );
spreadsheet.write( workbook, "c:/temp/data.xls" );
</cfscript>
```
### init()

When instantiating the library, the `init()` method **must** be called. This will happen automatically if you use the `New` keyword:
```
spreadsheet = New spreadsheetCFML.Spreadsheet();
```
If using `CreateObject()` then you must call `init()` explicitly:
```
spreadsheet = CreateObject( "component", "spreadsheetCFML.Spreadsheet" ).init();
```
### Using a mapping

You may wish to place the spreadsheet library files in a central location with an application mapping, and instantiate the component using its dot path (e.g. `New myLibrary.spreadsheetCFML.Spreadsheet();`).

[How to create mappings (StackOverflow)](http://stackoverflow.com/questions/12073714/components-mapping-in-railo).

[Full function reference](https://github.com/cfsimplicity/spreadsheet-cfml/wiki)

### ColdFusion spreadsheet functionality support

The library supports all of Adobe ColdFusion's spreadsheet functionality with a similar syntax. It can be run alongside existing ColdFusion spreadsheet code you don't wish to modify.

* [ColdFusion Spreadsheet functionality support](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/ColdFusion-spreadsheet-functionality-support)
)
* [Extra functions not available in ColdFusion](https://github.com/cfsimplicity/spreadsheet-cfml/wiki/Functions-not-available-in-ColdFusion)

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
spreadsheet = New spreadsheetCFML.Spreadsheet( dateFormats={ DATE: "mm/dd/yyyy" } );
</cfscript>
```
Or by using the `setDateFormats()` method on an existing instance.
```
<cfscript>
spreadsheet = New spreadsheetCFML.Spreadsheet();
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

Copyright (c) 2015-2025 Julian Halliwell

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
