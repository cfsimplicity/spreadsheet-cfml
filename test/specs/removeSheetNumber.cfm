<cfscript>
describe( "removeSheetNumber", function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Deletes the sheet number specified", function(){
		s.createSheet( workbook, "test" );
		s.removeSheetNumber( workbook, 2 );
		expect( workbook.getNumberOfSheets() ).toBe( 1 );
	});


	describe( "removeSheetNumber throws an exception if", function(){

		it( "the sheet number doesn't exist", function(){
			expect( function(){
				s.removeSheetNumber( workbook, 20 );
			}).toThrow( regex="Invalid sheet" );
		});

	});	

});	
</cfscript>