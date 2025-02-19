<cfscript>
describe( "newXlsx", ()=>{

	it( "Returns an XSSF workbook", ()=>{
		var workbook = s.newXlsx();
		expect( s.isXmlFormat( workbook ) ).toBeTrue();
	})

	it( "Creates a workbook with the specified sheet name", ()=>{
		var workbook = s.newXlsx( "test" );
		expect( s.getSheetHelper().getActiveSheetName( workbook ) ).toBe( "test" );
	})

})	
</cfscript>