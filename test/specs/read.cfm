<cfscript>
describe( "read", function(){

	beforeEach( function(){
		Sleep( 5 );// allow time for file operations to complete
	});

	it( "Can read a traditional XLS file", function(){
		var path = getTestFilePath( "test.xls" );
		var workbook = s.read( src=path );
		expect( s.isBinaryFormat( workbook ) ).toBeTrue();
	});

	it( "Can read an OOXML file", function(){
		var path = getTestFilePath( "test.xlsx" );
		var workbook = s.read( src=path );
		expect( s.isXmlFormat( workbook ) ).toBeTrue();
	});

	it( "Can read a traditional XLS file into a query", function(){
		var path = getTestFilePath( "test.xls" );
		var expected = querySim(
			"column1,column2
			a|b
			1|#ParseDateTime( '2015-04-01 00:00:00' )#
			#ParseDateTime( '2015-04-01 01:01:01' )#|2");
		var actual = s.read( src=path,format="query" );
		expect( actual ).toBe( expected );

	});

	it( "Can read an OOXML file into a query", function(){
		var path = getTestFilePath( "test.xlsx" );
		var expected = querySim(
			"column1,column2
			a|e
			b|f
			c|g
			I am|ooxml");
		var actual = s.read( src=path, format="query" );
	});

	it( "Reads from the specified sheet name", function(){
		var path = getTestFilePath( "test.xls" );// has 2 sheets
		var expected = querySim(
			"column1,column2
			x|y");
		var actual = s.read( src=path, format="query", sheetName="sheet2" );
		expect( actual ).toBe( expected );
	});

	it( "Reads from the specified sheet number", function(){
		var path = getTestFilePath( "test.xls" );// has 2 sheets
		var expected = querySim(
			"column1,column2
			x|y");
		var actual = s.read( src=path, format="query", sheetNumber=2 );
		expect( actual ).toBe( expected );
	});

	it( "Uses header row for column names if specified", function(){
		var path = getTestFilePath( "test.xls" );
		var expected = querySim(
			"a,b
			1|#ParseDateTime( '2015-04-01 00:00:00' )#
			#ParseDateTime( '2015-04-01 01:01:01' )#|2");
		var actual = s.read( src=path, format="query", headerRow=1 );
		expect( actual ).toBe( expected );
	});

	it( "Includes the specified header row in query if includeHeader is true", function(){
		var path = getTestFilePath( "test.xls" );
		var expected = querySim(
			"a,b
			a|b
			1|#ParseDateTime( '2015-04-01 00:00:00' )#
			#ParseDateTime( '2015-04-01 01:01:01' )#|2");
		var actual = s.read( src=path, format="query", headerRow=1, includeHeaderRow=true );
		expect( actual ).toBe( expected );
	});

	it( "Excludes null and blank rows in query by default", function(){
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "" ], [ "a", "b" ] ] );
		var workbook = s.new();
		s.addRows( workbook, data )
			.write( workbook, tempXlsPath, true );
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ] ] );;
		var actual = s.read( src=tempXlsPath, format="query" );
		expect( actual ).toBe( expected );
	});

	it( "Includes null and blank rows in query if includeBlankRows is true", function(){
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "" ], [ "a", "b" ] ] );
		var workbook = s.new();
		s.addRows( workbook, data )
			.write( workbook, tempXlsPath, true );
		var expected = data;
		var actual = s.read( src=tempXlsPath, format="query", includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	it( "Can handle null/empty cells", function(){
		var path = getTestFilePath( "nullCell.xls" );
		var actual = s.read( src=path, format="query", headerRow=1 );
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "a" ] ] );
		expect( actual ).toBe( expected );
	});

	it( "Writes and reads numeric, boolean, date and leading zero values correctly", function(){
		var dateValue = CreateDate( 2015, 04, 12 );
		var data = QueryNew( "column1,column2,column3,column4,column5", "Integer,Integer,Bit,Date,VarChar", [ [ 2, 0, true, dateValue, "01" ] ] );
		var workbook = s.new();
		s.addRows( workbook,data )
			.write( workbook, tempXlsPath, true );
		var expected = data;
		var actual = s.getSheetHelper().sheetToQuery( workbook );
		expect( actual ).toBe( expected );
		expect( IsNumeric( s.getCellValue( workbook, 1, 1 ) ) ).tobeTrue();
		expect( s.getCellValue( workbook, 1, 2 ) ).tobe( 0 );
		expect( IsBoolean( s.getCellValue( workbook, 1, 3 ) ) ).tobeTrue();
		expect( IsDate( s.getCellValue( workbook, 1, 4 ) ) ).tobeTrue();
	});

	it( "Can fill each of the empty cells in merged regions with the visible merged cell value without conflicting with includeBlankRows=true", function(){
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ], [ "", "" ] ] );
		var workbook = s.workbookFromQuery( data, false );
		s.mergeCells( workbook, 1, 2, 1, 2, true )//force empty merged cells
			.write( workbook, tempXlsPath, true );
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "a" ], [ "a", "a" ] ] );
		var actual = s.read( src=tempXlsPath, format="query", fillMergedCellsWithVisibleValue=true );
		expect( actual ).toBe( expected );
		//test retention of blank row not part of merge region
		expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "a" ], [ "a", "a" ], [ "", "" ] ] );
		actual = s.read( src=tempXlsPath, format="query", fillMergedCellsWithVisibleValue=true, includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	it( "Can read specified rows only into a query", function(){
		var data = QuerySim( "A
			A1
			A2
			A3
			A4
			A5");
		var workbook = s.workbookFromQuery( data, false );
		s.write( workbook, tempXlsPath, true );
		var actual = s.read( src=tempXlsPath, format="query", rows="2,4-5" );
		var expected = QuerySim( "column1
			A2
			A4
			A5");
		expect( actual ).toBe( expected );
		//with header row included in row 1
		data = QuerySim( "A1
			A2
			A3
			A4
			A5
			A6");
		var workbook = s.workbookFromQuery( data, true );
		s.write( workbook, tempXlsPath, true );
		var actual = s.read( src=tempXlsPath, format="query", rows="2,4-5", headerRow=1 );
		expected = QuerySim( "A1
			A2
			A4
			A5");
		expect( actual ).toBe( expected );
	});

	it( "Can read specified column numbers only into a query", function(){
		var data = QuerySim( "A,B,C,D,E
			A1|B1|C1|D1|E1");
		//With no header row, so no column names specified
		var workbook = s.workbookFromQuery( data, false );
		s.write( workbook, tempXlsPath,true );
		var actual = s.read( src=tempXlsPath, format="query", columns="2,4-5" );
		var expected = QuerySim( "column1,column2,column3
			B1|D1|E1");
		expect( actual ).toBe( expected );
		//With column names specified from the header row
		var workbook = s.workbookFromQuery( data, true );
		s.write( workbook ,tempXlsPath,true );
		actual = s.read( src=tempXlsPath, format="query", columns="2,4-5", headerRow=1 );
		expected = QuerySim( "B,D,E
			B1|D1|E1");
		expect( actual ).toBe( expected );
		//Include the header row with specified column names
		var workbook = s.workbookFromQuery( data, true );
		s.write( workbook, tempXlsPath, true );
		actual = s.read( src=tempXlsPath, format="query", columns="2,4-5", headerRow=1, includeHeaderRow=true );
		expected = QuerySim( "B,D,E
			B|D|E
			B1|D1|E1");
		expect( actual ).toBe( expected );
	});

	it( "Can read specific rows and columns only into a query", function(){
		var data = QuerySim( "A1,B1,C1,D1,E1
			A2|B2|C2|D2|E2
			A3|B3|C3|D3|E3
			A4|B4|C4|D4|E4
			A5|B5|C5|D5|E5");
		//First row is header
		var workbook = s.workbookFromQuery( data, true );
		s.write( workbook, tempXlsPath, true );
		var actual = s.read( src=tempXlsPath, format="query", columns="2,4-5", rows="2,4-5", headerRow=1 );
		var expected = QuerySim( "B1,D1,E1
			B2|D2|E2
			B4|D4|E4
			B5|D5|E5");
		expect( actual ).toBe( expected );
	});

	it( "Can return HTML table rows from an Excel file", function(){
		var path = getTestFilePath( "test.xls" );
		var actual = s.read( src=path, format="html" );
		var expected = "<tbody><tr><td>a</td><td>b</td></tr><tr><td>1</td><td>2015-04-01 00:00:00</td></tr><tr><td>2015-04-01 01:01:01</td><td>2</td></tr></tbody>";
		expect( actual ).toBe( expected );
		actual = s.read( src=path, format="html", headerRow=1 );
		expected = "<tbody><tr><td>1</td><td>2015-04-01 00:00:00</td></tr><tr><td>2015-04-01 01:01:01</td><td>2</td></tr></tbody>";
		expect( actual ).toBe( expected );
		actual = s.read( src=path, format="html", headerRow=1, includeHeaderRow=true );
		expected="<thead><tr><th>a</th><th>b</th></tr></thead><tbody><tr><td>a</td><td>b</td></tr><tr><td>1</td><td>2015-04-01 00:00:00</td></tr><tr><td>2015-04-01 01:01:01</td><td>2</td></tr></tbody>";
		expect( actual ).toBe( expected );
	});

	it( "Can return a CSV string from an Excel file", function(){
		var path = getTestFilePath( "test.xls" );
		var expected = 'a,b#crlf#1,2015-04-01 00:00:00#crlf#2015-04-01 01:01:01,2';
		var actual = s.read( src=path,format="csv" );
		expect( actual ).toBe( expected );
		expected = 'a,b#crlf#a,b#crlf#1,2015-04-01 00:00:00#crlf#2015-04-01 01:01:01,2';
		actual = s.read( src=path, format="csv", headerRow=1, includeHeaderRow=true );
		expect( actual ).toBe( expected );
	});

	it( "Escapes double-quotes in string values when reading to CSV", function(){
		var data = QueryNew( "column1", "VarChar", [ [ 'a "so-called" test' ] ] );
		var workbook = s.new();
		s.addRows( workbook, data )
			.write( workbook, tempXlsPath, true );
		var expected = '"a ""so-called"" test"';
		var actual = s.read( src=tempXlsPath, format="csv" );
		expect( actual ).toBe( expected );
	});

	it( "Accepts a custom delimiter when generating CSV", function(){
		var path = getTestFilePath( "test.xls" );
		var expected = 'a|b#crlf#1|2015-04-01 00:00:00#crlf#2015-04-01 01:01:01|2';
		var actual = s.read( src=path, format="csv", csvDelimiter="|" );
		expect( actual ).toBe( expected );
	});

	it( "Can exclude columns formatted as 'hidden'", function(){
		var workbook = s.new();
		s.addColumn( workbook, "a1" )
			.addColumn( workbook, "b1" )
			.hideColumn( workbook, 1 )
			.write( workbook, tempXlsPath, true );
		var actual = s.read( src=tempXlsPath, format="query", includeHiddenColumns=false );
		var expected = QuerySim( "column2
			b1");
		expect( actual ).toBe( expected );
	});

	it( "Returns an empty query if excluding hidden columns and ALL columns are hidden", function(){
		var workbook = s.new();
		s.addColumn( workbook, "a1" )
			.addColumn( workbook, "b1" )
			.hideColumn( workbook, 1 )
			.hideColumn( workbook, 2 )
			.write( workbook, tempXlsPath, true );
		var actual = s.read( src=tempXlsPath, format="query", includeHiddenColumns=false );
		var expected = QueryNew( "" );
		expect( actual ).toBe( expected );
	});

	it( "Can read an encrypted XLSX file", function(){
		var path = getTestFilePath( "passworded.xlsx" );
		var expected = QueryNew( "column1", "VarChar", [ [ "secret" ] ] );
		var actual = s.read( src=path, format="query", password="pass" );
		expect( actual ).toBe( expected );
	});

	it( "Can read an encrypted binary file", function(){
		var path = getTestFilePath( "passworded.xls" );
		var expected = QueryNew( "column1", "VarChar", [ [ "secret" ] ] );
		var actual = s.read( src=path, format="query", password="pass" );
		expect( actual ).toBe( expected );
	});
	
	it( "Can read a spreadsheet containing a formula", function(){
		var workbook = s.new();
		s.addColumn( workbook,"1,1" );
		var theFormula = "SUM(A1:A2)";
		s.setCellFormula( workbook, theFormula, 3, 1 )
			.write( workbook=workbook, filepath=tempXlsPath, overwrite=true );
		var expected = QueryNew( "column1","Integer", [ [ 1 ], [ 1 ], [ 2 ] ] );
		var actual = s.read( src=tempXlsPath, format="query" );
		expect( actual ).toBe( expected );
	});

	describe( "query column name setting", function() {

		it( "Allows column names to be specified as a list when reading a sheet into a query", function(){
			var path = getTestFilePath( "test.xls" );
			actual = s.read( src=path, format="query", columnNames="One,Two" );
			expected = QuerySim( "One,Two
				a|b
				1|#ParseDateTime( '2015-04-01 00:00:00' )#
				#ParseDateTime( '2015-04-01 01:01:01' )#|2");
			expect( actual ).toBe( expected );
		});

		it( "Allows column names to be specified as an array when reading a sheet into a query", function(){
			var path = getTestFilePath( "test.xls" );
			actual = s.read( src=path, format="query", columnNames=[ "One", "Two" ] );
			expected = QuerySim( "One,Two
				a|b
				1|#ParseDateTime( '2015-04-01 00:00:00' )#
				#ParseDateTime( '2015-04-01 01:01:01' )#|2");
			expect( actual ).toBe( expected );
		});

		it( "ColumnNames list overrides headerRow: none of the header row values will be used", function(){
			var path = getTestFilePath( "test.xls" );
			var actual = s.read( src=path, format="query", columnNames="One,Two", headerRow=1 );
			var expected = QuerySim( "One,Two
				1|#ParseDateTime( '2015-04-01 00:00:00' )#
				#ParseDateTime( '2015-04-01 01:01:01' )#|2");
			expect( actual ).toBe( expected );
		});

		it( "can handle column names containing commas or spaces", function(){
			var path = getTestFilePath( "commaAndSpaceInColumnHeaders.xls" );
			var actual = s.read( src=path, format="query", headerRow=1 );
			var columnNames = [ "first name", "surname,comma" ];// these are the file column headers
			expect( actual.getColumnNames()[ 1 ] ).toBe( columnNames[ 1 ] );
			expect( actual.getColumnNames()[ 2 ] ).toBe( columnNames[ 2 ] );
		});

		it( "Accepts 'queryColumnNames' as an alias of 'columnNames'", function(){
			var path = getTestFilePath( "test.xls" );
			actual = s.read( src=path, format="query", queryColumnNames="One,Two" );
			expected = QuerySim( "One,Two
				a|b
				1|#ParseDateTime( '2015-04-01 00:00:00' )#
				#ParseDateTime( '2015-04-01 01:01:01' )#|2");
			expect( actual ).toBe( expected );
		});

		it( "Allows header names to be made safe for query column names", function(){
			var data = [ [ "id","id","A  B","x/?y","(a)"," A","##1","1a" ], [ 1,2,3,4,5,6,7,8 ] ];
			var wb = s.newXlsx();
			s.addRows( wb, data )
				.write( wb, tempXlsxPath, true );
			var q = s.read( src=tempXlsxPath, format="query", headerRow=1, makeColumnNamesSafe=true );
			var expected = [ "id", "id2", "A_B", "x_y", "_a_", "A", "Number1", "_a" ];
			cfloop( from=1, to=expected.Len(), index="i" ){
				expect( q.getColumnNames()[ i ] ).toBe( expected[ i ] );
			}
			var wb = s.newXls();
			s.addRows( wb, data )
				.write( wb, tempXlsPath, true );
			var q = s.read( src=tempXlsPath, format="query", headerRow=1, makeColumnNamesSafe=true );
			var expected = [ "id", "id2", "A_B", "x_y", "_a_", "A", "Number1", "_a" ];
			cfloop( from=1, to=expected.Len(), index="i" ){
				expect( q.getColumnNames()[ i ] ).toBe( expected[ i ] );
			}
		});
		
	});

	describe( "query column type setting", function(){

		it( "allows the query column types to be manually set using list", function(){
			var workbook = s.new();
			s.addRow( workbook, [ 1, 1.1, "string", CreateTime( 1, 0, 0 ) ] )
				.write( workbook, tempXlsPath, true );
			var q = s.read( src=tempXlsPath, format="query", queryColumnTypes="Integer,Double,VarChar,Time" );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "INTEGER" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIME" );
		});

		it( "allows the query column types to be manually set where the column order isn't known, but the header row values are", function(){
			var workbook = s.new();
			s.addRows( workbook, [ [ "integer", "double", "string column", "time" ], [ 1, 1.1, "text", CreateTime( 1, 0, 0 ) ] ] )
				.write( workbook, tempXlsPath, true );
			var columnTypes = { "string column": "VARCHAR", "integer": "INTEGER", "time": "TIME", "double": "DOUBLE" };//not in order
			var q = s.read( src=tempXlsPath, format="query", queryColumnTypes=columnTypes, headerRow=1 );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "INTEGER" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIME" );
		});

		it( "allows the query column types to be manually set where the column order isn't known, but the column names are", function(){
			var workbook = s.new();
			s.addRows( workbook, [ [ 1, 1.1, "text", CreateTime( 1, 0, 0 ) ] ] )
				.write( workbook, tempXlsPath, true );
			var columnNames = "integer,double,string column,time";
			var columnTypes = { "string": "VARCHAR", "integer": "INTEGER", "time": "TIME", "double": "DOUBLE" };//not in order
			var q = s.read( src=tempXlsPath, format="query", queryColumnTypes=columnTypes, columnNames=columnNames );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "INTEGER" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIME" );
		});

		it( "allows the query column types to be automatically set", function(){
			var workbook = s.new();
			s.addRow( workbook, [ 1, 1.1, "string", Now() ] )
				.write( workbook, tempXlsPath, true );
			var q = s.read( src=tempXlsPath, format="query", queryColumnTypes="auto" );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIMESTAMP" );
		});

		it( "automatic detecting of query column types ignores blank cells", function(){
			var workbook = s.new();
			var data = [
				[ "", "", "", "" ],
				[ "", 2, "test", Now() ],
				[ 1, 1.1, "string", Now() ],
				[ 1, "", "", "" ]
			];
			s.addRows( workbook, data )
				.write( workbook, tempXlsPath, true );
			var q = s.read( src=tempXlsPath, format="query", queryColumnTypes="auto" );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIMESTAMP" );
		});

		it( "allows a default type to be set for all query columns", function(){
			var workbook = s.new();
			s.addRow( workbook, [ 1, 1.1, "string", Now() ] )
				.write( workbook, tempXlsPath, true );
			var q = s.read( src=tempXlsPath, format="query", queryColumnTypes="VARCHAR" );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 2 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "VARCHAR" );
		});

	});

	describe( "read throws an exception if", function(){

		it( "queryColumnTypes is specified as a 'columnName/type' struct, but headerRow and columnNames arguments are missing", function(){
			expect( function(){
				var columnTypes = { col1: "Integer" };
				s.read( src=getTestFilePath( "test.xlsx" ), format="query", queryColumnTypes=columnTypes );
			}).toThrow( regex="Invalid argument" );
		});

		it( "a formula can't be evaluated", function(){
			expect( function(){
				var workbook = s.new();
				s.addColumn( workbook, "1,1" );
				var theFormula="SUS(A1:A2)";//invalid formula
				s.setCellFormula( workbook, theFormula, 3, 1 )
					.write( workbook=workbook, filepath=tempXlsPath, overwrite=true )
					.read( src=tempXlsPath, format="query" );
			}).toThrow( regex="Failed to run formula" );
		});

		it( "the 'query' argument is passed", function(){
			expect( function(){
				s.read( src=tempXlsPath, query="q" );
			}).toThrow( regex="Invalid argument" );
		});

		it( "the format argument is invalid", function(){
			expect( function(){
				s.read( src=tempXlsPath, format="wrong" );
			}).toThrow( regex="Invalid format" );
		});

		it( "the file doesn't exist", function(){
			expect( function(){
				var path = getTestFilePath( "nonexistent.xls" );
				s.read( src=path );
			}).toThrow( regex="Non-existent file" );
		});

		it( "the sheet name doesn't exist", function(){
			expect( function(){
				var path = getTestFilePath( "test.xls" );
				s.read( src=path, format="query", sheetName="nonexistent" );
			}).toThrow( regex="Invalid sheet" );
		});

		it( "the sheet number doesn't exist", function(){
			expect( function(){
				var path = getTestFilePath( "test.xls" );
				s.read( src=path, format="query", sheetNumber=20 );
			}).toThrow( regex="Invalid sheet|out of range" );
		});

		it( "the password for an encrypted XML file is incorrect", function(){
			expect( function(){
				var tempXlsxPath = getTestFilePath( "passworded.xlsx" );
				s.read( src=tempXlsxPath, format="query", password="parse" );
			}).toThrow( regex="(Invalid password|Password incorrect|password is invalid)" );
		});

		it( "the password for an encrypted binary file is incorrect", function(){
			expect( function(){
				var xlsPath = getTestFilePath( "passworded.xls" );
				s.read( src=xlsPath, format="query", password="parse" );
			}).toThrow( regex="(Invalid password|Password incorrect|password is invalid)" );
		});

		it( "the source file is not a spreadsheet", function(){
			expect( function(){
				var path = getTestFilePath( "notaspreadsheet.txt" );
				s.read( src=path );
			}).toThrow( regex="Invalid spreadsheet file" );
		});

		it( "the source file appears to contain CSV or TSV, and suggests using 'csvToQuery'", function(){
			expect( function(){
				var path = getTestFilePath( "csv.xls" );
				s.read( src=path );
			}).toThrow( regex="may be a CSV" );
			expect( function(){
				var path = getTestFilePath( "test.tsv" );
				s.read( src=path );
			}).toThrow( regex="may be a CSV" );
		});

		it( "the source file is in an old format not supported by POI", function(){
			expect( function(){
				var path = getTestFilePath( "oldformat.xls" );
				s.read( src=path );
			}).toThrow( regex="Invalid spreadsheet format" );
		});

	});

	afterEach( function(){
		if( FileExists( variables.tempXlsPath ) ) FileDelete( variables.tempXlsPath );
		if( FileExists( variables.tempXlsxPath ) ) FileDelete( variables.tempXlsxPath );
	});

});	
</cfscript>