<cfscript>
describe( "chaining", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Allows void methods to be chained", function() {
		var theComment = {
			author: "cfsimplicity"
			,comment: "This is the comment in row 1 column 1"
		};
		var expected = Duplicate( theComment ).Append( { column: 1, row: 1 } );
		workbooks.Each( function( wb ){
			s.addColumn( wb, "1" ).setCellComment( wb, theComment, 1, 1 );
			var actual = s.getCellComment( wb, 1, 1 );
			expect( actual ).toBe( expected );
		});
	});

});	
</cfscript>