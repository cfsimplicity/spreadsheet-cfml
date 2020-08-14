<cfscript>
describe( "addColumn",function(){

	beforeEach( function(){
		variables.columnData = "a,b";
		variables.dataAsArray = [ "a", "b" ];
		variables.workbook = s.new();
	});

	it( "Adds a column with the minimum arguments",function() {
		s.addColumn( workbook,columnData );
		expected = QueryNew( "column1","VarChar",[ [ "a" ],[ "b" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Adds a column with the minimum arguments using array data",function() {
		s.addColumn( workbook, dataAsArray );
		expected = QueryNew( "column1", "VarChar", [ [ "a" ], [ "b" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Adds a column at a given start row",function() {
		s.addColumn( workbook,columnData,2 );
		expected = QueryNew( "column1","VarChar",[ [ "" ],[ "a" ],[ "b" ] ] );
		actual = s.sheetToQuery( workbook=workbook,includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	it( "Adds a column at a given column number",function() {
		s.addColumn( workbook=workbook,data=columnData,startColumn=2 );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "","a" ],[ "","b" ] ] );
		actual = s.sheetToQuery( workbook=workbook,includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	it( "Adds a column including commas with a custom delimiter",function() {
		var columnData = "a,b|c,d";
		s.addColumn( workbook=workbook,data=columnData,delimiter="|" );
		expected = QueryNew( "column1","VarChar",[ [ "a,b" ],[ "c,d" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Inserts (not replaces) a column with the minimum arguments",function() {
		s.addColumn( workbook,columnData );
		s.addColumn( workbook=workbook,data=columnData,insert=true );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","a" ],[ "b","b" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Adds numeric values correctly",function() {
		var rowData = "1,1.1";
		s.addColumn( workbook, rowData );
		expect( s.getCellValue( workbook, 1, 1 ) ).toBe( 1 );
		expect( s.getCellValue( workbook, 2, 1 ) ).toBe( 1.1 );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "numeric" );
		expect( s.getCellType( workbook, 2, 1 ) ).toBe( "numeric" );
	});

it( "Adds boolean values as strings",function() {
		var rowData = true;
		s.addColumn( workbook, rowData );
		expect( s.getCellValue( workbook, 1, 1 ) ).toBe( true );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "string" );
	});

	it( "Adds date/time values correctly",function() {
		var dateValue = CreateDate( 2015, 04, 12 );
		var timeValue = CreateTime( 1, 0, 0 );
		var dateTimeValue = CreateDateTime( 2015, 04, 12, 1, 0, 0 );
		var rowData = "#dateValue#,#timeValue#,#dateTimeValue#";
		s.addColumn( workbook, rowData );
		expect( s.getCellValue( workbook, 1, 1 ) ).toBe( dateValue );
		expect( s.getCellValue( workbook, 2, 1 ) ).toBe( timeValue );
		expect( s.getCellValue( workbook, 3, 1 ) ).toBe( dateTimeValue );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "numeric" );
		expect( s.getCellType( workbook, 2, 1 ) ).toBe( "numeric" );
		expect( s.getCellType( workbook, 3, 1 ) ).toBe( "numeric" );
	});

	it( "Adds zeros as zeros, not booleans",function(){
		s.addColumn( workbook, 0 );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "numeric" );
	});

	it( "Adds strings with leading zeros as strings not numbers",function(){
		s.addColumn( workbook, "01" );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "string" );
	});

});	
</cfscript>