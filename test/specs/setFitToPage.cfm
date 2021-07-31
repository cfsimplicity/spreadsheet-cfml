<cfscript>
describe( "setFitToPage", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "sets the active sheet's print setup to fit everything in one page by default", function(){
		workbooks.Each( function( wb ){
			var sheet = s.getSheetHelper().getActiveSheet( wb );
			s.setFitToPage( wb, false );
			expect( sheet.getFitToPage() ).toBeFalse();
			s.setFitToPage( wb, true );
			expect( sheet.getFitToPage() ).toBeTrue();
			expect( sheet.getPrintSetup().getFitWidth() ).toBe( 1 );
			expect( sheet.getPrintSetup().getFitHeight() ).toBe( 1 );
		});
	});

	it( "allows the number of pages wide and high to be specified", function(){
		workbooks.Each( function( wb ){
			var sheet = s.getSheetHelper().getActiveSheet( wb );
			s.setFitToPage( wb, true, 2, 0 );
			expect( sheet.getFitToPage() ).toBeTrue();
			expect( sheet.getPrintSetup().getFitWidth() ).toBe( 2 );
			expect( sheet.getPrintSetup().getFitHeight() ).toBe( 0 );
		});
	});

});	
</cfscript>