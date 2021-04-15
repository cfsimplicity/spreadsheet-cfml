<cfscript>
describe( "readBinary", function(){

	it( "Returns a binary object", function(){
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			expect( IsBinary( s.readBinary( wb ) ) ).toBeTrue();
		});
	});

});	
</cfscript>