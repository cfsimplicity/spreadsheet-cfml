<cfscript>
describe( "binaryFromQuery tests",function(){

	it( "Returns a binary object",function() {
		data = QueryNew( "Header1,Header2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
		expect( IsBinary( s.binaryFromQuery( data ) ) ).toBeTrue();
	});

});	
</cfscript>