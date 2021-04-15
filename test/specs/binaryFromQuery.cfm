<cfscript>
describe( "binaryFromQuery", function(){

	it( "Returns a binary object", function() {
		var data = QueryNew( "Header1,Header2", "VarChar,VarChar", [ [ "a","b" ], [ "c","d" ] ] );
		expect( IsBinary( s.binaryFromQuery( data ) ) ).toBeTrue();
		expect( IsBinary( s.binaryFromQuery( data=data, xmlFormat=true ) ) ).toBeTrue();
	});

});	
</cfscript>