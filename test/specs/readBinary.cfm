<cfscript>
describe( "readBinary",function(){

	it( "Returns a binary object",function() {
		var workbook = s.new();
		expect( IsBinary( s.readBinary( workbook ) ) ).toBeTrue();
	});

});	
</cfscript>