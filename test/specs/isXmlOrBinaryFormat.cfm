<cfscript>
describe( "isXmlOrBinaryFormat", function(){

	it( "Reports a binary file's format correctly", function(){
		var path = getTestFilePath( "test.xls" );//binary file
		var workbook = s.read( src=path );
		expect( s.isBinaryFormat( workbook ) ).toBeTrue();
		expect( s.isXmlFormat( workbook ) ).toBeFalse();
		expect( s.isStreamingXmlFormat( workbook ) ).toBeFalse();
	});

	it( "Reports an XML file's format correctly", function(){
		var path = getTestFilePath( "test.xlsx" );//binary file
		var workbook = s.read( src=path );
		expect( s.isXmlFormat( workbook ) ).toBeTrue();
		expect( s.isBinaryFormat( workbook ) ).toBeFalse();
		expect( s.isStreamingXmlFormat( workbook ) ).toBeFalse();
	});

	it( "Reports a streaming XML file's format correctly", function(){
		var workbook = s.newStreamingXlsx();
		expect( s.isStreamingXmlFormat( workbook ) ).toBeTrue();
		expect( s.isXmlFormat( workbook ) ).toBeTrue();
		expect( s.isBinaryFormat( workbook ) ).toBeFalse();
	});

});	
</cfscript>
