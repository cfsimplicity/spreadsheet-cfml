<cfscript>
describe( "isSpreadsheetFile tests",function(){

	it( "reports false for a non-spreadsheet file",function() {
		path=ExpandPath( "/root/test/files/notaspreadsheet.txt" );
		expect( s.isSpreadsheetFile( path ) ).toBeFalse();
	});

	it( "reports true for a binary spreadsheet file",function() {
		path=ExpandPath( "/root/test/files/test.xls" );
		expect( s.isSpreadsheetFile( path ) ).toBeTrue();
	});

	it( "reports true for an xml spreadsheet file",function() {
		path=ExpandPath( "/root/test/files/test.xlsx" );
		expect( s.isSpreadsheetFile( path ) ).toBeTrue();
	});

	describe( "isSpreadsheetFile exceptions",function(){
		it( "Throws an exception if the file doesn't exist",function() {
			expect( function(){
				var path=ExpandPath( "/root/test/files/nonexistant.xls" );
				s.isSpreadsheetFile( path )
			}).toThrow( regex="Non-existent file" );
		});
	});

});	
</cfscript>
