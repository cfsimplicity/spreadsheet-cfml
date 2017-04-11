<cfscript>
describe( "isSpreadsheetFile",function(){

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

	describe( "Throws an exception if",function(){

		it( "the file doesn't exist",function() {
			expect( function(){
				var path=ExpandPath( "/root/test/files/nonexistant.xls" );
				s.isSpreadsheetFile( path );
			}).toThrow( regex="Non-existent file" );
		});
		
	});

});	
</cfscript>