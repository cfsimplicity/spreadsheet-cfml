<cfscript>
describe( "setCellValue tests",function(){

	beforeEach( function(){
		variables.workbook = s.new();
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