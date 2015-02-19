<cfscript>
describe( "setCellValue tests",function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Sets the specified cell to the specified value",function() {
		data =  QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
		s.addRows( workbook,data );
		expected = "d";
		actual = s.getCellValue( workbook,2,2 );
		expect( actual ).toBe( expected );
	});

});	
</cfscript>