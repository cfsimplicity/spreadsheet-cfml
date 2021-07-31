<cfscript>
describe( "newXls", function(){

	it( "Returns an HSSF workbook", function(){
		var workbook = s.newXls();
		expect( s.isBinaryFormat( workbook ) ).toBeTrue();
	});

	it( "Creates a workbook with the specified sheet name", function(){
		var workbook = s.newXls( "test" );
		expect( s.getSheetHelper().getActiveSheetName( workbook ) ).toBe( "test" );
	});

});	
</cfscript>