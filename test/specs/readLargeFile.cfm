<cfscript>
describe(
	title="readLargeFile (Lucee only)"
	,body=function(){

	it( "Can read an XLSX file into a query", function(){
		var path = getTestFilePath( "large.xlsx" );
		var expected = querySim(
			"column1,column2
			FirstSheet A1|FirstSheet B1");
		var actual = s.readLargeFile( src=path );
	});

	it( "Reads from the specified sheet name", function(){
		var path = getTestFilePath( "large.xlsx" );// has 2 sheets
		var expected = querySim(
			"column1,column2
			SecondSheet A1|SecondSheet B1");
		var actual = s.readLargeFile( src=path, sheetName="SecondSheet" );
		expect( actual ).toBe( expected );
	});

	it( "Reads from the specified sheet name", function(){
		var path = getTestFilePath( "large.xlsx" );// has 2 sheets
		var expected = querySim(
			"column1,column2
			SecondSheet A1|SecondSheet B1");
		var actual = s.readLargeFile( src=path, sheetNumber=2 );
		expect( actual ).toBe( expected );
	});

	it( "Uses the specifed header row for column names", function(){
		var path = getTestFilePath( "large.xlsx" );
		var expected = querySim(
			"heading1,heading2
			A2 value|B2 value");
		var actual = s.readLargeFile( src=path, headerRow=1, sheetName="HeaderRow" );
		expect( actual ).toBe( expected );
	});

	it( "Generates default column names if the data has more columns than the specifed header row", function(){
		var headerRow = [ "firstColumn" ];
		var dataRow1 = [ "row 1 col 1 value" ];
		var dataRow2 = [ "row 2 col 1 value", "row 2 col 2 value" ];
		var expected = querySim(
			"firstColumn,column2
			row 1 col 1 value|
			row 2 col 1 value|row 2 col 2 value"
		);
		s.newChainable( "xlsx" )
		 .addRow( headerRow )
		 .addRow( dataRow1 )
		 .addRow( dataRow2 )
		 .write( tempXlsxPath, true );
		var actual = s.readLargeFile( src=tempXlsxPath, headerRow=1 );
		expect( actual ).toBe( expected );
	});

	it( "Includes the specified header row in query if includeHeader is true", function(){
		var headerRow = [ "a", "b" ];
		var dataRow = [ "c", "d" ];
		s.newChainable( "xlsx" )
		 .addRow( headerRow )
		 .addRow( dataRow )
		 .write( tempXlsxPath, true );
		var expected = querySim(
			"a,b
			a|b
			c|d");
		var actual = s.readLargeFile( src=tempXlsxPath, headerRow=1, includeHeaderRow=true );
		expect( actual ).toBe( expected );
	});

	it( "Excludes null and blank rows in query by default", function(){
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "" ], [ "a", "b" ] ] );
		var workbook = s.newXlsx();
		s.addRows( workbook, data )
			.write( workbook, tempXlsxPath, true );
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ] ] );
		var actual = s.readLargeFile( src=tempXlsxPath );
		expect( actual ).toBe( expected );
	});

	it( "Includes null and blank rows in query if includeBlankRows is true", function(){
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "" ], [ "a", "b" ] ] );
		var workbook = s.newXlsx();
		s.addRows( workbook, data )
			.write( workbook, tempXlsxPath, true );
		var expected = data;
		var actual = s.readLargeFile( src=tempXlsxPath, includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	it( "Can handle null/empty cells", function(){
		var path = getTestFilePath( "nullCell.xlsx" );
		var actual = s.readLargeFile( src=path, headerRow=1 );
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "a" ] ] );
		expect( actual ).toBe( expected );
	});

	it( "Includes trailing empty columns when using a header row", function(){
		var expected = QuerySim( "col1,col2,emptyCol
			value|value|");
		var workbook = s.newChainable( "xlsx" )
			.addRow( "col1,col2,emptyCol" )
			.addRow( "value,value" )
			.write( tempXlsxPath, true );
		var actual = s.readLargeFile( src=tempXlsxPath, headerRow=1 );
		expect( actual ).toBe( expected );
	});

	it( "Can return HTML table rows from an Excel file", function(){
		var headerRow = [ "header1", "header2" ];
		var dataRow = [ "a", CreateDate( 2015, 04, 01 ) ];
		s.newChainable( "xlsx" )
		 .addRow( headerRow )
		 .addRow( dataRow )
		 .write( tempXlsxPath, true );
		var actual = s.readLargeFile( src=tempXlsxPath, format="html" );
		var expected = "<tbody><tr><td>header1</td><td>header2</td></tr><tr><td>a</td><td>2015-04-01 00:00:00</td></tr></tbody>";
		expect( actual ).toBe( expected );
		actual = s.readLargeFile( src=tempXlsxPath, format="html", headerRow=1 );
		expected = "<tbody><tr><td>a</td><td>2015-04-01 00:00:00</td></tr></tbody>";
		expect( actual ).toBe( expected );
		actual = s.read( src=tempXlsxPath, format="html", headerRow=1, includeHeaderRow=true );
		expected = "<thead><tr><th>header1</th><th>header2</th></tr></thead><tbody><tr><td>header1</td><td>header2</td></tr><tr><td>a</td><td>2015-04-01 00:00:00</td></tr></tbody>";
		expect( actual ).toBe( expected );
	});

	it( "Can return a CSV string from an Excel file", function(){
		var headerRow = [ "header1", "header2" ];
		var dataRow = [ "a", CreateDate( 2015, 04, 01 ) ];
		s.newChainable( "xlsx" )
		 .addRow( headerRow )
		 .addRow( dataRow )
		 .write( tempXlsxPath, true );
		var expected = 'header1,header2#newline#a,2015-04-01 00:00:00#newline#';
		var actual = s.readLargeFile( src=tempXlsxPath, format="csv" );
		expect( actual ).toBe( expected );
		expected = 'header1,header2#newline#header1,header2#newline#a,2015-04-01 00:00:00#newline#';
		actual = s.readLargeFile( src=tempXlsxPath, format="csv", headerRow=1, includeHeaderRow=true );
		expect( actual ).toBe( expected );
	});

	it( "Accepts a custom delimiter when generating CSV", function(){
		var dataRow = [ "a", CreateDate( 2015, 04, 01 ) ];
		s.newChainable( "xlsx" )
		 .addRow( dataRow )
		 .write( tempXlsxPath, true );
		var expected = 'a|2015-04-01 00:00:00#newline#';
		var actual = s.readLargeFile( src=tempXlsxPath, format="csv", csvDelimiter="|" );
		expect( actual ).toBe( expected );
	});

	it( "Includes columns formatted as 'hidden' by default", function(){
		s.newChainable( "xlsx" )
			.addColumn( "a1" )
			.addColumn( "b1" )
			.hideColumn( 1 )
			.write( tempXlsPath, true );
		var actual = s.readLargeFile( src=tempXlsPath );
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a1", "b1" ] ] );
		expect( actual ).toBe( expected );
	});

	it( "Can exclude columns formatted as 'hidden'", function(){
		s.newChainable( "xlsx" )
			.addColumn( "a1" )
			.addColumn( "b1" )
			.hideColumn( 1 )
			.write( tempXlsPath, true );
		var actual = s.readLargeFile( src=tempXlsPath, includeHiddenColumns=false );
		var expected = QueryNew( "column2", "VarChar", [ [ "b1" ] ] );
		expect( actual ).toBe( expected );
	});

	it( "Includes rows formatted as 'hidden' by default", function(){
		var data = QueryNew( "column1", "VarChar", [ [ "Apple" ], [ "Banana" ], [ "Carrot" ], [ "Doughnut" ] ] );
		s.newChainable( "xlsx" )
			.addRows( data )
			.hideRow( 1 )
			.write( tempXlsPath, true );
		var actual = s.readLargeFile( src=tempXlsPath );
		var expected = data;
		expect( actual ).toBe( expected );
	});

	it( "Can exclude rows formatted as 'hidden'", function(){
		var data = QueryNew( "column1", "VarChar", [ [ "Apple" ], [ "Banana" ], [ "Carrot" ], [ "Doughnut" ] ] );
		s.newChainable( "xlsx" )
			.addRows( data )
			.hideRow( 1 )
			.write( tempXlsPath, true );
		var actual = s.readLargeFile( src=tempXlsPath, includeHiddenRows=false );
		var expected = QueryNew( "column1", "VarChar", [ [ "Banana" ], [ "Carrot" ], [ "Doughnut" ] ] );
		expect( actual ).toBe( expected );
	});

	it( "Returns an empty query if the spreadsheet is empty even if headerRow is specified", function(){
		s.newChainable( "xlsx" )
			.write( tempXlsPath, true );
		var actual = s.readLargeFile( src=tempXlsPath, headerRow=1 );
		var expected = QueryNew( "" );
		expect( actual ).toBe( expected );
	});

	it( "Returns an empty query if excluding hidden columns and ALL columns are hidden", function(){
		s.newChainable( "xlsx" )
			.addColumn( "a1" )
			.addColumn( "b1" )
			.hideColumn( 1 )
			.hideColumn( 2 )
			.write( tempXlsPath, true );
		var actual = s.readLargeFile( src=tempXlsPath, includeHiddenColumns=false );
		var expected = QueryNew( "" );
		expect( actual ).toBe( expected );
	});

	it( "Returns a query with column names but no rows if column names are specified but spreadsheet is empty", function(){
		s.newChainable( "xlsx" ).write( tempXlsxPath, true );
		var actual = s.readLargeFile( src=tempXlsxPath, queryColumnNames="One,Two" );
		var expected = QueryNew( "One,Two","Varchar,Varchar", [] );
		expect( actual ).toBe( expected );
	});

	it( "Can read an encrypted XLSX file", function(){
		var path = getTestFilePath( "passworded.xlsx" );
		var expected = QueryNew( "column1", "VarChar", [ [ "secret" ] ] );
		var actual = s.readLargeFile( src=path, password="pass" );
		expect( actual ).toBe( expected );
	});

	it( "Can read a spreadsheet containing a CACHED (i.e. pre-evaluated) formula", function(){
		/* NB: Setting a formula with POI does not cache its value. The Streaming Reader cannot evaluate formulas */
		var path = getTestFilePath( "formula.xlsx" );
		var expected = QueryNew( "column1","Integer", [ [ 1 ], [ 1 ], [ 2 ] ] );
		var actual = s.readLargeFile( path );
		expect( actual ).toBe( expected );
	});

	it( "Returns raw cell values by default", function() {
		var rawValue = 0.000011;
		s.newChainable( "xlsx" )
			.setCellValue( rawValue, 1, 1, "numeric" )
			.formatCell( { dataformat: "0.00000" }, 1, 1 )
			.write( tempXlsxPath, true );
		var actual = s.readLargeFile( tempXlsxPath );
		expect( actual.column1 ).toBe( rawValue );
	});

	it( "Can return visible/formatted values rather than raw values", function() {
		var rawValue = 0.000011;
		var visibleValue = 0.00001;
		s.newChainable( "xlsx" )
			.setCellValue( rawValue, 1, 1, "numeric" )
			.formatCell( { dataformat: "0.00000" }, 1, 1 )
			.write( tempXlsxPath, true );
		var actual = s.readLargeFile( src=tempXlsxPath, returnVisibleValues=true );
		expect( actual.column1 ).toBe( visibleValue );
	});

	describe( "query column name setting", function() {

		it( "Allows column names to be specified as a list when reading a sheet into a query", function(){
			s.newChainable( "xlsx" ).addRow( "a,b" ).write( tempXlsxPath, true );
			var actual = s.readLargeFile( src=tempXlsxPath, queryColumnNames="One,Two" );
			var expected = QueryNew( "One,Two","Varchar,Varchar", [ "a", "b" ] );
			expect( actual ).toBe( expected );
		});

		it( "Allows column names to be specified as an array when reading a sheet into a query", function(){
			s.newChainable( "xlsx" ).addRow( "a,b" ).write( tempXlsxPath, true );
			var actual = s.readLargeFile( src=tempXlsxPath, queryColumnNames=[ "One", "Two" ] );
			var expected = QueryNew( "One,Two","Varchar,Varchar", [ "a", "b" ] );
			expect( actual ).toBe( expected );
		});

		it( "ColumnNames list overrides headerRow: none of the header row values will be used", function(){
			s.newChainable( "xlsx" ).addRow( "a,b" ).addRow( "c,d" ).write( tempXlsxPath, true );
			var actual = s.readLargeFile( src=tempXlsxPath, queryColumnNames="One,Two", headerRow=1 );
			var expected = QueryNew( "One,Two","Varchar,Varchar", [ "c", "d" ] );
			expect( actual ).toBe( expected );
		});

		it( "can handle column names containing commas or spaces", function(){
			var path = getTestFilePath( "commaAndSpaceInColumnHeaders.xlsx" );
			var actual = s.readLargeFile( src=path, headerRow=1 );
			var columnNames = [ "first name", "surname,comma" ];// these are the file column headers
			expect( actual.getColumnNames()[ 1 ] ).toBe( columnNames[ 1 ] );
			expect( actual.getColumnNames()[ 2 ] ).toBe( columnNames[ 2 ] );
		});

		it( "Allows header names to be made safe for query column names", function(){
			var data = [ [ "id","id","A  B","x/?y","(a)"," A","##1","1a" ], [ 1,2,3,4,5,6,7,8 ] ];
			s.newChainable( "xlsx" ).addRows( data ).write( tempXlsxPath, true );
			var q = s.readLargeFile( src=tempXlsxPath, headerRow=1, makeColumnNamesSafe=true );
			var expected = [ "id", "id2", "A_B", "x_y", "_a_", "A", "Number1", "_a" ];
			cfloop( from=1, to=expected.Len(), index="i" ){
				expect( q.getColumnNames()[ i ] ).toBe( expected[ i ] );
			}
		});

		it( "Generates default column names if the data has more columns than the specifed column names", function(){
			var columnNames = [ "firstColumn" ];
			var dataRow1 = [ "row 1 col 1 value" ];
			var dataRow2 = [ "row 2 col 1 value", "row 2 col 2 value" ];
			var expected = querySim(
				"firstColumn,column2
				row 1 col 1 value|
				row 2 col 1 value|row 2 col 2 value"
			);
			s.newChainable( "xlsx" ).addRow( dataRow1 ).addRow( dataRow2 ).write( tempXlsxPath, true );
			var actual = s.readLargeFile( src=tempXlsxPath, queryColumnNames=columnNames );
			expect( actual ).toBe( expected );
		});

	});

	describe( "query column type setting", function(){

		it( "allows the query column types to be manually set using list", function(){
			s.newChainable( "xlsx" ).addRow( [ 1, 1.1, "string", _CreateTime( 1, 0, 0 ) ] ).write( tempXlsxPath, true );
			var q = s.readLargeFile( src=tempXlsxPath, queryColumnTypes="Integer,Double,VarChar,Time" );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "INTEGER" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIME" );
		});

		it( "allows the query column types to be manually set where the column order isn't known, but the header row values are", function(){
			s.newChainable( "xlsx" )
				.addRows( [ [ "integer", "double", "string column", "time" ], [ 1, 1.1, "text", _CreateTime( 1, 0, 0 ) ] ] )
				.write( tempXlsxPath, true );
			var columnTypes = { "string column": "VARCHAR", "integer": "INTEGER", "time": "TIME", "double": "DOUBLE" };//not in order
			var q = s.readLargeFile( src=tempXlsxPath, format="query", queryColumnTypes=columnTypes, headerRow=1 );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "INTEGER" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIME" );
		});

		it( "allows the query column types to be manually set where the column order isn't known, but the column names are", function(){
			s.newChainable( "xlsx" ).addRows( [ [ 1, 1.1, "text", _CreateTime( 1, 0, 0 ) ] ] ).write( tempXlsxPath, true );
			var columnNames = "integer,double,string column,time";
			var columnTypes = { "string": "VARCHAR", "integer": "INTEGER", "time": "TIME", "double": "DOUBLE" };//not in order
			var q = s.readLargeFile( src=tempXlsxPath, queryColumnTypes=columnTypes, queryColumnNames=columnNames );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "INTEGER" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIME" );
		});

		it( "allows the query column types to be automatically set", function(){
			s.newChainable( "xlsx" ).addRow( [ 1, 1.1, "string", Now() ] ).write( tempXlsxPath, true );
			var q = s.readLargeFile( src=tempXlsxPath, queryColumnTypes="auto" );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIMESTAMP" );
		});

		it( "automatic detecting of query column types ignores blank cells", function(){
			var data = [
				[ "", "", "", "" ],
				[ "", 2, "test", Now() ],
				[ 1, 1.1, "string", Now() ],
				[ 1, "", "", "" ]
			];
			s.newChainable( "xlsx" ).addRows( data ).write( tempXlsxPath, true );
			var q = s.readLargeFile( src=tempXlsxPath, queryColumnTypes="auto" );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIMESTAMP" );
		});

		it( "allows a default type to be set for all query columns", function(){
			s.newChainable( "xlsx" ).addRow( [ 1, 1.1, "string", Now() ] ).write( tempXlsxPath, true );
			var q = s.readLargeFile( src=tempXlsxPath, queryColumnTypes="VARCHAR" );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 2 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "VARCHAR" );
		});

	});

	describe( "readLargeFile throws an exception if", function(){

		it( "the file doesn't exist", function(){
			expect( function(){
				var path = getTestFilePath( "nonexistent.xls" );
				s.readLargeFile( src=path );
			}).toThrow( type="cfsimplicity.spreadsheet.nonExistentFile" );
		});
		
		it( "the file to be read is not an XLSX type", function(){
			expect( function(){
				var path = getTestFilePath( "test.xls" );
				s.readLargeFile( src=path );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidSpreadsheetType" );
		});

		it( "both sheetName and sheetNumber arguments are specified", function(){
			expect( function(){
				var path = getTestFilePath( "large.xlsx" );
				s.readLargeFile( src=path, sheetName="sheet1", sheetNumber=2 );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidArguments" );
		});

		it( "the format argument is invalid", function(){
			expect( function(){
				s.readLargeFile( src=getTestFilePath( "large.xlsx" ), format="wrong" );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidReadFormat" );
		});

		it( "the sheet name doesn't exist", function(){
			expect( function(){
				s.readLargeFile( src=getTestFilePath( "large.xlsx" ), sheetName="nonexistent" );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetName" );
		});

		it( "the sheet number doesn't exist", function(){
			expect( function(){
				s.readLargeFile( src=getTestFilePath( "large.xlsx" ), sheetNumber=20 );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetNumber" );
		});

		it( "the source file is not a spreadsheet", function(){
			expect( function(){
				s.readLargeFile( src=getTestFilePath( "notaspreadsheet.txt" ) );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidSpreadsheetType" );
		});

	});

	describe( "the streaming reader", function(){

		it( "allows options to be passed", function(){
			var options = {
				bufferSize: 512
				,rowCacheSize: 5
			};
			var builder = s.getStreamingReaderHelper().getBuilder( options );
			expect( builder.getBufferSize() ).toBe( options.bufferSize );
			expect( builder.getRowCacheSize() ).toBe( options.rowCacheSize );
		});

	});

	afterEach( function(){
		if( FileExists( variables.tempXlsxPath ) )
			FileDelete( variables.tempXlsxPath );
	});

	}
	,skip=s.getIsACF()
);

describe(
	title="readLargeFile (when run on ACF)"
	,body=function(){

		it( "throws a methodNotSupported exception", function(){
			expect( function(){
				var path = getTestFilePath( "large.xlsx" );
				s.readLargeFile( src=path );
			}).toThrow( type="cfsimplicity.spreadsheet.methodNotSupported" );
		});

	}
	,skip=!s.getIsACF()
); 
</cfscript>