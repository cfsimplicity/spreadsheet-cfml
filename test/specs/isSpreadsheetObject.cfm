<cfscript>
describe( "isSpreadsheetObject",function(){

	it( "reports false for a variable which is not a spreadsheet object",function() {
		var objectToTest="a string";
		expect( s.isSpreadsheetObject( objectToTest ) ).toBeFalse();
	});

	it( "reports true for a binary spreadsheet object",function() {
		var path=ExpandPath( "/root/test/files/test.xls" );
		var objectToTest=s.read( path );
		expect( s.isSpreadsheetObject( objectToTest ) ).toBeTrue();
	});

	it( "reports true for an xml spreadsheet object",function() {
		var path=ExpandPath( "/root/test/files/test.xlsx" );
		var objectToTest=s.read( path );
		expect( s.isSpreadsheetObject( objectToTest ) ).toBeTrue();
	});

});	
</cfscript>
