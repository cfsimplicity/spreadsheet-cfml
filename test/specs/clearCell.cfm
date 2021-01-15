<cfscript>
describe( "clearCell", function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Clears the specified cell", function(){
		s.setCellValue( workbook, "test", 1, 1 );
		s.clearCell( workbook, 1, 1 );
		var expected = "";
		var actual = s.getCellValue( workbook, 1, 1 );
		expect( actual ).toBe( expected );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "BLANK" );
	});

	it( "Clears the specified range of cells", function(){
		var data = QueryNew( "column1,column2,column3", "VarChar,VarChar,VarChar", [ [ "a","b","c" ], [ "d","e","f" ], [ "g","h","i" ] ] );
		s.addRows( workbook,data );
		s.clearCellRange( workbook,2,2,3,3 );
		var expected = QueryNew( "column1,column2,column3", "VarChar,VarChar,VarChar", [ [ "a","b","c" ], [ "d","","" ], [ "g","","" ] ] );
		var actual = s.sheetToQuery( workbook=workbook, includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

});	
</cfscript>