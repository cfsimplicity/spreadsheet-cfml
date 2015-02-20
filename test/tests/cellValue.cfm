<cfscript>
describe( "cellValue tests",function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Gets the value from the specified cell",function() {
		data =  QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
		s.addRows( workbook,data );
		expect( s.getCellValue( workbook,2,2 ) ).toBe( "d" );
	});

	it( "Sets the specified cell to the specified value",function() {
		value="test";
		s.setCellValue( workbook,value,1,1 );
		expected = querySim( "column1
			test");
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

});	
</cfscript>