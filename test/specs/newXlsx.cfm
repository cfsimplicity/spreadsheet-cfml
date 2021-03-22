<cfscript>
describe( "newXlsx", function(){

	it( "Returns an XSSF workbook", function(){
		var workbook = s.newXlsx();
		expect( s.isXmlFormat( workbook ) ).toBeTrue();
	});

	it( "Creates a workbook with the specified sheet name", function(){
		var workbook = s.newXlsx( "test" );
		makePublic( s, "getActiveSheetName" );
		expect( s.getActiveSheetName( workbook ) ).toBe( "test" );
	});

});	
</cfscript>