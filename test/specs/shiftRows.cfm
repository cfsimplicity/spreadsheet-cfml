<cfscript>
describe( "shiftRows tests",function(){

	beforeEach( function(){
		variables.data = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
		variables.workbook = s.new();
	});

	it( "Shifts rows down if offset is positive",function() {
		s.addRows( workbook,data );
		s.shiftRows( workbook,1,1,1 )
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "","" ],[ "a","b" ] ] );
		actual = s.sheetToQuery( workbook=workbook,includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	it( "Shifts rows up if offset is negative",function() {
		s.addRows( workbook,data );
		s.shiftRows( workbook,2,2,-1)
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "c","d" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

});	
</cfscript>