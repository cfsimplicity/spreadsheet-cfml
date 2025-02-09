<cfscript>
describe( "new", ()=>{

	it( "Returns an HSSF workbook by default", ()=>{
		var workbook = s.new();
		expect( s.isBinaryFormat( workbook ) ).toBeTrue();
	})

	it( "Returns an XSSF workbook if xmlFormat is true", ()=>{
		var workbook = s.new( xmlFormat=true );
		expect( s.isXmlFormat( workbook ) ).toBeTrue();
	})

	it( "Returns an XSSF workbook if global defaultWorkbookFormat is 'xml' or 'xlsx'", ()=>{
		var s = newSpreadsheetInstance().setDefaultWorkbookFormat( "xml" );
		var workbook = s.new();
		expect( s.isXmlFormat( workbook ) ).toBeTrue();
		s.setDefaultWorkbookFormat( "binary" );
		workbook = s.new();
		expect( s.isBinaryFormat( workbook ) ).toBeTrue();
		s.setDefaultWorkbookFormat( "xlsx" );
		workbook = s.new();
		expect( s.isXmlFormat( workbook ) ).toBeTrue();
	})

	it( "Returns a streaming XSSF workbook if streamingXml is true", ()=>{
		var workbook = s.new( streamingXml=true );
		expect( s.isStreamingXmlFormat( workbook ) ).toBeTrue();
	})

	it( "Creates a workbook with the specified sheet name", ()=>{
		var workbook = s.new( "test" );
		expect( s.getSheetHelper().getActiveSheetName( workbook ) ).toBe( "test" );
	})

	describe( "new throws an exception if", ()=>{

		it( "the sheet name contains invalid characters", ()=>{
			expect( ()=>{
				s.new( "[]?*\/:" );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidCharacters" );
		})

		it( "streaming XML is specified with an invalid streamingWindowSize", ()=>{
			expect( ()=>{
				s.new( xmlFormat=true, streamingXml=true, streamingWindowSize=1.2 );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidStreamingWindowSizeArgument" );
			expect( ()=>{
				s.new( xmlFormat=true, streamingXml=true, streamingWindowSize=-1 );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidStreamingWindowSizeArgument" );
		})

	})

})	
</cfscript>