<cfscript>
describe( "addColumn",function(){

	beforeEach( function(){
		variables.columnData = "a,b";
		variables.workbook = s.new();
	});

	it( "Adds a column with the minimum arguments",function() {
		s.addColumn( workbook,columnData );
		expected = QueryNew( "column1","VarChar",[ [ "a" ],[ "b" ] ] );
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

	it( "Adds numeric, boolean or date values correctly",function() {
		var dateValue = CreateDate( 2015, 04, 12 );
		s.addColumn( workbook, "2" );
		s.addColumn( workbook=workbook, data=true, startColumn=2 );
		s.addColumn( workbook=workbook, data=dateValue, startColumn=3 );
		expected = QueryNew( "column1,column2,column3", "Integer,Bit,Date", [ [ 2, true, dateValue ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
		expect( s.getCellType( workbook, 1 , 1 ) ).toBe( "numeric" );
		expect( s.getCellType( workbook, 1 , 2 ) ).toBe( "string" );
		expect( s.getCellType( workbook, 1 , 3 ) ).toBe( "numeric" );
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