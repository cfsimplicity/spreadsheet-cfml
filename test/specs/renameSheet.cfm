<cfscript>
describe( "renameSheet", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Renames the specified sheet", function(){
		workbooks.Each( function( wb ){
			s.renameSheet( wb, "test", 1 );
			expect( s.getSheetHelper().sheetExists( wb, "test" ) ).toBeTrue();
		});
	});


	describe( "renameSheet throws an exception if", function(){

		it( "the new sheet name already exists", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.createSheet( wb, "test" )
						.createSheet( wb, "test2" )
						.renameSheet( wb, "test2", 2 );
				}).toThrow( regex="Invalid Sheet Name" );
			});
		});

	});	

});	
</cfscript>