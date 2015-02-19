<cfscript>
describe( "addColumn tests",function(){

	beforeEach( function(){
		variables.data = "a,b";
		variables.workbook = s.new();
	});

	it( "Adds a column with the minimum arguments",function() {
		s.addColumn( workbook,data );
		expected = QueryNew( "column1","VarChar",[ [ "a" ],[ "b" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Adds a column at a given start row",function() {
		s.addColumn( workbook,data,2 );
		expected = QueryNew( "column1","VarChar",[ [ "" ],[ "a" ],[ "b" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Adds a column at a given column number",function() {
		s.addColumn( workbook=workbook,data=data,startColumn=2 );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "","a" ],[ "","b" ] ] );
		actual = s.sheetToQuery( workbook=workbook,includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	it( "Adds a column including commas with a custom delimiter",function() {
		var data = "a,b|c,d";
		s.addColumn( workbook=workbook,data=data,delimiter="|" );
		expected = QueryNew( "column1","VarChar",[ [ "a,b" ],[ "c,d" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Inserts (not replaces) a column with the minimum arguments",function() {
		s.addColumn( workbook,data );
		s.addColumn( workbook=workbook,data=data,insert=true );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","a" ],[ "b","b" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

});	
</cfscript>
