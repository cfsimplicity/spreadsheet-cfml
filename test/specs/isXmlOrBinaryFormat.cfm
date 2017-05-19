<cfscript>
describe( "isXmlOrBinaryFormat",function(){

	it( "Reports a binary file's format correctly",function() {
		path = getTestFilePath( "test.xls" );//binary file
		workbook = s.read( src=path );
		expect( s.isBinaryFormat( workbook ) ).toBeTrue();
		expect( s.isXmlFormat( workbook ) ).toBeFalse();
	});

	it( "Reports an XML file's format correctly",function() {
		path = getTestFilePath( "test.xlsx" );//binary file
		workbook = s.read( src=path );
		expect( s.isXmlFormat( workbook ) ).toBeTrue();
		expect( s.isBinaryFormat( workbook ) ).toBeFalse();
	});

});	
</cfscript>
