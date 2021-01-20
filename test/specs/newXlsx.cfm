<cfscript>
describe( "newXlsx", function(){

	it( "Returns an XSSF workbook", function(){
		var workbook = s.newXlsx();
		expect( workbook.getClass().name ).toBe( "org.apache.poi.xssf.usermodel.XSSFWorkbook" );
	});

	it( "Creates a workbook with the specified sheet name", function(){
		var workbook = s.newXlsx( "test" );
		makePublic( s,"getActiveSheetName" );
		expect( s.getActiveSheetName( workbook ) ).toBe( "test" );
	});

});	
</cfscript>