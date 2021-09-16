<cfscript>
describe( "readBinary", function(){

	it( "Returns a binary object", function(){
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			expect( IsBinary( s.readBinary( wb ) ) ).toBeTrue();
		});
	});

	it( "Is chainable", function(){
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			var actual = s.newChainable( wb ).readBinary();
			expect( IsBinary( actual ) ).toBeTrue();
		});
	});

});	
</cfscript>