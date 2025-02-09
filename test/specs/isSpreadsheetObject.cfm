<cfscript>
describe( "isSpreadsheetObject", ()=>{

	it( "reports false for a variable which is not a spreadsheet object", ()=>{
		var objectToTest = "a string";
		expect( s.isSpreadsheetObject( objectToTest ) ).toBeFalse();
	})

	it( "reports true for a binary spreadsheet object", ()=>{
		var path = getTestFilePath( "test.xls" );
		var objectToTest = s.read( path );
		expect( s.isSpreadsheetObject( objectToTest ) ).toBeTrue();
	})

	it( "reports true for an xml spreadsheet object", ()=>{
		var path = getTestFilePath( "test.xlsx" );
		var objectToTest = s.read( path );
		expect( s.isSpreadsheetObject( objectToTest ) ).toBeTrue();
	})

})	
</cfscript>
