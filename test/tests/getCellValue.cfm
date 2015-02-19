<cfscript>
describe( "getCellValue tests",function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Gets the value from the specified cell",function() {
		value="test";
		s.setCellValue( workbook,value,1,1 );
		expected = querySim( "column1
			test");
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

});	
</cfscript>