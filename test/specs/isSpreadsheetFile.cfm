<cfscript>
describe( "isSpreadsheetFile", ()=>{

	it( "reports false for a non-spreadsheet file", ()=>{
		var path = getTestFilePath( "notaspreadsheet.txt" );
		expect( s.isSpreadsheetFile( path ) ).toBeFalse();
	})

	it( "reports true for a binary spreadsheet file", ()=>{
		var path = getTestFilePath( "test.xls" );
		expect( s.isSpreadsheetFile( path ) ).toBeTrue();
	})

	it( "reports true for an xml spreadsheet file", ()=>{
		var path = getTestFilePath( "test.xlsx" );
		expect( s.isSpreadsheetFile( path ) ).toBeTrue();
	})

	describe( "isSpreadsheetFile throws an exception if", ()=>{

		it( "the file doesn't exist", ()=>{
			expect( ()=>{
				var path = getTestFilePath( "nonexistent.xls" );
				s.isSpreadsheetFile( path );
			}).toThrow( type="cfsimplicity.spreadsheet.nonExistentFile" );
		})

		it( "the source file is in an old format not supported by POI", ()=>{
			expect( ()=>{
				var path = getTestFilePath( "oldformat.xls" );
				s.isSpreadsheetFile( path );
			}).toThrow( type="cfsimplicity.spreadsheet.oldExcelFormatException" );
		})
		
	})

})	
</cfscript>