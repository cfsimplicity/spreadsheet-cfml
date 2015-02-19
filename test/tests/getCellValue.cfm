<cfscript>
describe( "getCellValue tests",function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Gets the value from the specified cell",function() {
		data =  QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
		s.addRows( workbook,data );
		expect( s.getCellValue( workbook,2,2 ) ).toBe( "d" );
	});

});	
</cfscript>