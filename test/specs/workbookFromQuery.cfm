<cfscript>
describe( "workbookFromQuery",function(){

	beforeEach( function(){
		variables.query = QueryNew( "Header1,Header2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
		variables.workbook = s.new();
	});

	it( "Returns a workbook from a query",function() {
		workbook = s.workbookFromQuery( query );
		expected = query;
		actual = s.sheetToQuery( workbook=workbook,headerRow=1 );
		expect( actual ).toBe( expected );
	});

	it( "Returns an XSSF workbook if xmlFormat is true",function() {
		workbook = s.workbookFromQuery( data=query,xmlformat=true );
		expect( workbook.getClass().name ).toBe( "org.apache.poi.xssf.usermodel.XSSFWorkbook" );
	});

	it( "Adds the header row in the same case and order as the query columns", function() {
		local.query = QueryNew( "Header2,Header1", "VarChar,VarChar", [ [ "b","a" ], [ "d","c" ] ] );
		workbook = s.workbookFromQuery( data=local.query, addHeaderRow=true );
		expect( s.getCellValue( workbook, 1, 1 ) ).toBeWithCase( "Header2" );
	});

});	
</cfscript>