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
			c|d");
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
		expect( actual ).toBe( expected );
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
			c|d");
		actual = s.read( src=path,format="query",headerRow=1 );
		expect( actual ).toBe( expected );
	});

	it( "Includes the specified header row in query if includeHeader is true",function() {
		path = ExpandPath( "/root/test/files/test.xls" );
		expected = querySim(
			"a,b
			a|b
			c|d");
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