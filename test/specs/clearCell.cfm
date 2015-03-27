<cfscript>
describe( "clearCell tests",function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Clears the specified cell",function() {
		s.setCellValue( workbook,value,1,1 );
		s.clearCell( workbook,1,1 );
		expected = "";
		actual = s.getCellValue( workbook,1,1 );
		expect( actual ).toBe( expected );
	});

	it( "Clears the specified range of cells",function() {
		data = QueryNew( "column1,column2,column3","VarChar,VarChar,VarChar",[ [ "a","b","c" ],[ "d","e","f" ],[ "g","h","i" ] ] );
		s.addRows( workbook,data );
		s.clearCellRange( workbook,2,2,3,3 );
		expected = QueryNew( "column1,column2,column3","VarChar,VarChar,VarChar",[ [ "a","b","c" ],[ "d","","" ],[ "g","","" ] ] );
		actual = s.sheetToQuery( workbook=workbook,includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

});	
</cfscript>