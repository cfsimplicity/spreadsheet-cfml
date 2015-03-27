<cfscript>
describe( "workbookFromQuery tests",function(){

	beforeEach( function(){
		variables.data = QueryNew( "Header1,Header2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
		variables.workbook = s.new();
	});

	it( "Returns a workbook from a query",function() {
		workbook = s.workbookFromQuery( data );
		expected = data;
		actual = s.sheetToQuery( workbook=workbook,headerRow=1 );
		expect( actual ).toBe( expected );
	});

	it( "Returns an XSSF workbook if xmlFormat is true",function() {
		workbook = s.workbookFromQuery( data=data,xmlformat=true );
		expect( workbook.getClass().name ).toBe( "org.apache.poi.xssf.usermodel.XSSFWorkbook" );
	});

});	
</cfscript>