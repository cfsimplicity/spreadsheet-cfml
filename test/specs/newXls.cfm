<cfscript>
describe( "newXls", ()=>{

	it( "Returns an HSSF workbook", ()=>{
		var workbook = s.newXls();
		expect( s.isBinaryFormat( workbook ) ).toBeTrue();
	})

	it( "Creates a workbook with the specified sheet name", ()=>{
		var workbook = s.newXls( "test" );
		expect( s.getSheetHelper().getActiveSheetName( workbook ) ).toBe( "test" );
	})

})	
</cfscript>