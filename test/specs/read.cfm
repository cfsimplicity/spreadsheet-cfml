<cfscript>
describe( "read tests",function(){

	it( "Can read a traditional XLS file",function() {
		path = ExpandPath( "/root/test/files/test.xls" );
		workbook = s.read( src=path );
		expect( s.isBinaryFormat( workbook ) ).toBeTrue();
	});

	it( "Can read an OOXML file",function() {
		path = ExpandPath( "/root/test/files/test.xlsx" );
		workbook = s.read( src=path );
		expect( s.isXmlFormat( workbook ) ).toBeTrue();
	});

	it( "Can read a traditional XLS file into a query",function() {
		path = ExpandPath( "/root/test/files/test.xls" );
		expected = querySim(
			"column1,column2
			a|b
			1|#ParseDateTime( '2015-04-01 00:00:00' )#
			#ParseDateTime( '2015-04-01 01:01:01' )#|2");
		actual = s.read( src=path,format="query" );
		expect( actual ).toBe( expected );

	});

	it( "Can read an OOXML file into a query",function() {
		path = ExpandPath( "/root/test/files/test.xlsx" );
		expected = querySim(
			"column1,column2
			a|e
			b|f
			c|g
			I am|ooxml");
		actual = s.read( src=path,format="query" );
	});

	it( "Reads from the specified sheet name",function(){
		path = ExpandPath( "/root/test/files/test.xls" );// has 2 sheets
		expected = querySim(
			"column1,column2
			x|y");
		actual = s.read( src=path,format="query",sheetName="sheet2" );
		expect( actual ).toBe( expected );
	});

	it( "Reads from the specified sheet number",function(){
		path = ExpandPath( "/root/test/files/test.xls" );// has 2 sheets
		expected = querySim(
			"column1,column2
			x|y");
		actual = s.read( src=path,format="query",sheetNumber=2 );
		expect( actual ).toBe( expected );
	});

	it( "Uses header row for column names if specified",function() {
		path = ExpandPath( "/root/test/files/test.xls" );
		expected = querySim(
			"a,b
			1|#ParseDateTime( '2015-04-01 00:00:00' )#
			#ParseDateTime( '2015-04-01 01:01:01' )#|2");
		actual = s.read( src=path,format="query",headerRow=1 );
		expect( actual ).toBe( expected );
	});

	it( "Includes the specified header row in query if includeHeader is true",function() {
		path = ExpandPath( "/root/test/files/test.xls" );
		expected = querySim(
			"a,b
			a|b
			1|#ParseDateTime( '2015-04-01 00:00:00' )#
			#ParseDateTime( '2015-04-01 01:01:01' )#|2");
		actual = s.read( src=path,format="query",headerRow=1,includeHeaderRow=true );
		expect( actual ).toBe( expected );
	});

	it( "Excludes null and blank rows in query by default",function() {
		data = QueryNew( "column1,column2","VarChar,VarChar",[ [ "","" ],[ "a","b" ] ] );
		workbook = s.new();
		s.addRows( workbook,data );
		s.write( workbook,tempXlsPath,true );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","b" ] ] );;
		actual = s.read( src=tempXlsPath,format="query" );
		expect( actual ).toBe( expected );
	});

	it( "Includes null and blank rows in query if includeBlankRows is true",function() {
		data = QueryNew( "column1,column2","VarChar,VarChar",[ [ "","" ],[ "a","b" ] ] );
		workbook = s.new();
		s.addRows( workbook,data );
		s.write( workbook,tempXlsPath,true );
		expected = data;
		actual = s.read( src=tempXlsPath,format="query",includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	it( "Can handle null/empty cells",function() {
		path = ExpandPath( "/root/test/files/nullCell.xls" );
		actual = s.read( src=path ,format="query",headerRow=1 );
		expected=QueryNew( "column1,column2","VarChar,VarChar",[ [ "","a" ] ] );
		expect( actual ).toBe( expected );
	});

	it( "Writes and reads numeric, boolean and date values correctly",function() {
		var dateValue = CreateDate( 2015,04,12 );
		var data = QueryNew( "column1,column2,column3","Numeric,Boolean,Date",[ [ 2,true,dateValue ] ] );
		workbook = s.new();
		s.addRows( workbook,data );
		s.write( workbook,tempXlsPath,true );
		expected = data;
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
		expect( IsNumeric( s.getCellValue( workbook,1,1 ) ) ).tobeTrue();
		expect( IsBoolean( s.getCellValue( workbook,1,2 ) ) ).tobeTrue();
		expect( IsDate( s.getCellValue( workbook,1,3 ) ) ).tobeTrue();
	});

	it( "Can fill each of the empty cells in merged regions with the visible merged cell value without conflicting with includeBlankRows=true",function() {
		data = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ],[ "","" ] ] );
		workbook = s.workbookFromQuery( data,false );
		s.mergeCells( workbook,1,2,1,2,true );//force empty merged cells
		s.write( workbook,tempXlsPath,true );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","a" ],[ "a","a" ] ] );
		actual = s.read( src=tempXlsPath,format="query",fillMergedCellsWithVisibleValue=true );
		expect( actual ).toBe( expected );
		//test retention of blank row not part of merge region
		expected=QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","a" ],[ "a","a" ],[ "","" ] ] );
		actual = s.read( src=tempXlsPath,format="query",fillMergedCellsWithVisibleValue=true,includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	it( "Can read specified rows only into a query",function() {
		data=QuerySim( "A
			A1
			A2
			A3
			A4
			A5");
		workbook = s.workbookFromQuery( data,false );
		s.write( workbook,tempXlsPath,true );
		var actual	=	s.read( src=tempXlsPath,format="query",rows="2,4-5" );
		expected =	QuerySim( "column1
			A2
			A4
			A5");
		expect( actual ).toBe( expected );
		//with header row included in row 1
		data=QuerySim( "A1
			A2
			A3
			A4
			A5
			A6");
		workbook = s.workbookFromQuery( data,true );
		s.write( workbook,tempXlsPath,true );
		var actual	=	s.read( src=tempXlsPath,format="query",rows="2,4-5",headerRow=1 );
		expected =	QuerySim( "A1
			A2
			A4
			A5");
		expect( actual ).toBe( expected );
	}); 

	it( "Can read specified column numbers only into a query",function() {
		data=QuerySim( "A,B,C,D,E
			A1|B1|C1|D1|E1");
		//With no header row, so no column names specified
		workbook = s.workbookFromQuery( data,false );
		s.write( workbook,tempXlsPath,true );
		var actual	=	s.read( src=tempXlsPath,format="query",columns="2,4-5" );
		expected = QuerySim( "column1,column2,column3
			B1|D1|E1");
		expect( actual ).toBe( expected );
		//With column names specified from the header row
		workbook = s.workbookFromQuery( data,true );
		s.write( workbook,tempXlsPath,true );
		actual	=	s.read( src=tempXlsPath,format="query",columns="2,4-5",headerRow=1 );
		expected = QuerySim( "B,D,E
			B1|D1|E1");
		expect( actual ).toBe( expected );
		//Include the header row with specified column names
		workbook = s.workbookFromQuery( data,true );
		s.write( workbook,tempXlsPath,true );
		actual	=	s.read( src=tempXlsPath,format="query",columns="2,4-5",headerRow=1,includeHeaderRow=true );
		expected =	QuerySim( "B,D,E
			B|D|E
			B1|D1|E1");
		expect( actual ).toBe( expected );
	});

	it( "Can read specific rows and columns only into a query",function() {
		data=QuerySim( "A1,B1,C1,D1,E1
			A2|B2|C2|D2|E2
			A3|B3|C3|D3|E3
			A4|B4|C4|D4|E4
			A5|B5|C5|D5|E5");
		//First row is header
		workbook = s.workbookFromQuery( data,true );
		s.write( workbook,tempXlsPath,true );
		actual	=	s.read( src=tempXlsPath,format="query",columns="2,4-5",rows="2,4-5",headerRow=1 );
		expected = QuerySim( "B1,D1,E1
			B2|D2|E2
			B4|D4|E4
			B5|D5|E5");
		expect( actual ).toBe( expected );
	});

	it( "Allows column names to be specified as a list when reading a sheet into a query",function() {
		path = ExpandPath( "/root/test/files/test.xls" );
		// only one column name specified. The other will be the default
		actual = s.read( src=path,format="query",columnNames="One" );
		expected = QuerySim( "One,column2
			a|b
			1|#ParseDateTime( '2015-04-01 00:00:00' )#
			#ParseDateTime( '2015-04-01 01:01:01' )#|2");
		expect( actual ).toBe( expected );
		//both names specified
		actual = s.read( src=path,format="query",columnNames="One,Two" );
		expected = QuerySim( "One,Two
			a|b
			1|#ParseDateTime( '2015-04-01 00:00:00' )#
			#ParseDateTime( '2015-04-01 01:01:01' )#|2");
		expect( actual ).toBe( expected );
	});

	it( "ColumnNames list overrides headerRow: none of the header row values will be used",function() {
		path = ExpandPath( "/root/test/files/test.xls" );
		actual = s.read( src=path,format="query",columnNames="One,Two",headerRow=1 );
		expected = QuerySim( "One,Two
			1|#ParseDateTime( '2015-04-01 00:00:00' )#
			#ParseDateTime( '2015-04-01 01:01:01' )#|2");
		expect( actual ).toBe( expected );
	});

	it( "Can return HTML table rows from an Excel file",function() {
		path = ExpandPath( "/root/test/files/test.xls" );
		actual = s.read( src=path,format="html" );
		expected="<tbody><tr><td>a</td><td>b</td></tr><tr><td>1</td><td>2015-04-01 00:00:00</td></tr><tr><td>2015-04-01 01:01:01</td><td>2</td></tr></tbody>";
		expect( actual ).toBe( expected );
		actual = s.read( src=path,format="html",headerRow=1 );
		expected="<thead><tr><th>a</th><th>b</th></tr></thead><tbody><tr><td>1</td><td>2015-04-01 00:00:00</td></tr><tr><td>2015-04-01 01:01:01</td><td>2</td></tr></tbody>";
		expect( actual ).toBe( expected );
		actual = s.read( src=path,format="html",headerRow=1,includeHeaderRow=true );
		expected="<thead><tr><th>a</th><th>b</th></tr></thead><tbody><tr><td>a</td><td>b</td></tr><tr><td>1</td><td>2015-04-01 00:00:00</td></tr><tr><td>2015-04-01 01:01:01</td><td>2</td></tr></tbody>";
		expect( actual ).toBe( expected );
	});

	it( "Can return a CSV string from an Excel file",function() {
		path = ExpandPath( "/root/test/files/test.xls" );
		var crlf=Chr( 13 ) & Chr( 10 );
		expected='"a","b"#crlf#"1","2015-04-01 00:00:00"#crlf#"2015-04-01 01:01:01","2"';
		actual = s.read( src=path,format="csv" );
		expect( actual ).toBe( expected );
		expected='"a","b"#crlf#"a","b"#crlf#"1","2015-04-01 00:00:00"#crlf#"2015-04-01 01:01:01","2"';
		actual = s.read( src=path,format="csv",headerRow=1,includeHeaderRow=true );
		expect( actual ).toBe( expected );
	});

	it( "Escapes double-quotes in string values when reading to CSV",function() {
		data = QueryNew( "column1","VarChar",[ [ 'a "so-called" test' ] ] );
		workbook = s.new();
		s.addRows( workbook,data );
		s.write( workbook,tempXlsPath,true );
		expected = '"a ""so-called"" test"';
		actual = s.read( src=tempXlsPath,format="csv" );
		expect( actual ).toBe( expected );
	});

	it( "Can exclude columns formatted as 'hidden'",function() {
		workbook = s.new();
		s.addColumn( workbook,"a1" );
		s.addColumn( workbook,"b1" );
		s.hideColumn( workbook,1 );
		s.write( workbook,tempXlsPath,true );
		var actual	=	s.read( src=tempXlsPath,format="query",includeHiddenColumns=false );
		expected=QuerySim( "column2
			b1");
		expect( actual ).toBe( expected );
	});

	it( "Returns an empty query if excluding hidden columns and ALL columns are hidden",function() {
		workbook = s.new();
		s.addColumn( workbook,"a1" );
		s.addColumn( workbook,"b1" );
		s.hideColumn( workbook,1 );
		s.hideColumn( workbook,2 );
		s.write( workbook,tempXlsPath,true );
		var actual	=	s.read( src=tempXlsPath,format="query",includeHiddenColumns=false );
		expected=Query();
		expect( actual ).toBe( expected );
	});

	describe( "read exceptions",function(){

		it( "Throws an exception if the 'query' argument is passed",function() {
			expect( function(){
				s.read( src=tempXlsPath,query="q" );
			}).toThrow( regex="Invalid argument" );
		});

		it( "Throws an exception if the format argument is invalid",function() {
			expect( function(){
				s.read( src=tempXlsPath,format="wrong" );
			}).toThrow( regex="Invalid format" );
		});

		it( "Throws an exception if the sheet name doesn't exist",function() {
			expect( function(){
				s.read( src=path,format="query",sheetName="nonexistant" );
			}).toThrow( regex="Invalid sheet" );
		});

		it( "Throws an exception if the sheet number doesn't exist",function() {
			expect( function(){
				s.read( src=path,format="query",sheetNumber=20 );
			}).toThrow( regex="Invalid sheet|out of range" );
		});

	});

});	
</cfscript>