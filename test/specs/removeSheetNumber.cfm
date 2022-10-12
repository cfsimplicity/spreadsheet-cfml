<cfscript>
describe( "removeSheetNumber", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Deletes the sheet number specified", function(){
		workbooks.Each( function( wb ){
			s.createSheet( wb, "test" )
				.removeSheetNumber( wb, 2 );
			expect( wb.getNumberOfSheets() ).toBe( 1 );
		});
	});

	it( "Is chainable", function(){
		workbooks.Each( function( wb ){
			s.newChainable( wb )
				.createSheet( "test" )
				.removeSheetNumber( 2 );
			expect( wb.getNumberOfSheets() ).toBe( 1 );
		});
	});

	describe( "removeSheetNumber throws an exception if", function(){

		it( "the sheet number doesn't exist", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.removeSheetNumber( wb, 20 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetNumber" );
			});
		});

	});	

});	
</cfscript>