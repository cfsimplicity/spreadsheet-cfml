<cfscript>
describe( "new", function(){

	it( "Returns an HSSF workbook by default", function(){
		var workbook = s.new();
		expect( s.isBinaryFormat( workbook ) ).toBeTrue();
	});

	it( "Returns an XSSF workbook if xmlFormat is true", function(){
		var workbook = s.new( xmlFormat=true );
		expect( s.isXmlFormat( workbook ) ).toBeTrue();
	});

	it( "Returns a streaming XSSF workbook if streamingXml is true", function(){
		var workbook = s.new( streamingXml=true );
		expect( s.isStreamingXmlFormat( workbook ) ).toBeTrue();
	});

	it( "Creates a workbook with the specified sheet name", function(){
		var workbook = s.new( "test" );
		makePublic( s, "getActiveSheetName" );
		expect( s.getActiveSheetName( workbook ) ).toBe( "test" );
	});

	describe( "new throws an exception if", function(){

		it( "the sheet name contains invalid characters", function(){
			expect( function(){
				s.new( "[]?*\/:" );
			}).toThrow( regex="Invalid characters" );
		});

		it( "streaming XML is specified with an invalid streamingWindowSize", function(){
			expect( function(){
				s.new( xmlFormat=true, streamingXml=true, streamingWindowSize=1.2 );
			}).toThrow( regex="Invalid 'streamingWindowSize' argument" );
			expect( function(){
				s.new( xmlFormat=true, streamingXml=true, streamingWindowSize=-1 );
			}).toThrow( regex="Invalid 'streamingWindowSize' argument" );
		});

	});

});	
</cfscript>