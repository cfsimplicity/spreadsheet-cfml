<cfscript>
describe( "setActiveSheetNumber", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Sets the specified sheet number to be active", function(){
		workbooks.Each( function( wb ){
			s.createSheet( wb, "test" );
			s.setActiveSheetNumber( wb, 2 );
			expect( s.getSheetHelper().getActiveSheetName( wb ) ).toBe( "test" );
		});
	});


	describe( "setActiveSheetNumber throws an exception if", function(){

		it( "the sheet number doesn't exist", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.setActiveSheetNumber( wb, 20 );
				}).toThrow( regex="Invalid sheet" );
			});
		});

	});	

});	
</cfscript>