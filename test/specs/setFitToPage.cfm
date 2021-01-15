<cfscript>
describe( "setFitToPage", function(){

	beforeEach( function(){
		variables.xls = s.new();
		variables.xlsx = s.newXlsx();
	});

	it( "sets the active sheet's print setup to fit everything in one page by default", function(){
		makePublic( s, "getActiveSheet" );
		var sheet = s.getActiveSheet( xls );
		s.setFitToPage( xls, false );
		expect( sheet.getFitToPage() ).toBeFalse();
		s.setFitToPage( xls, true );
		expect( sheet.getFitToPage() ).toBeTrue();
		expect( sheet.getPrintSetup().getFitWidth() ).toBe( 1 );
		expect( sheet.getPrintSetup().getFitHeight() ).toBe( 1 );
		//xlsx
		sheet = s.getActiveSheet( xlsx );
		s.setFitToPage( xlsx, false );
		expect( sheet.getFitToPage() ).toBeFalse();
		s.setFitToPage( xlsx, true );
		expect( sheet.getFitToPage() ).toBeTrue();
		expect( sheet.getPrintSetup().getFitWidth() ).toBe( 1 );
		expect( sheet.getPrintSetup().getFitHeight() ).toBe( 1 );
	});

	it( "allows the number of pages wide and high to be specified", function(){
		makePublic( s, "getActiveSheet" );
		var sheet = s.getActiveSheet( xls );
		s.setFitToPage( xls, true, 2, 0 );
		expect( sheet.getFitToPage() ).toBeTrue();
		expect( sheet.getPrintSetup().getFitWidth() ).toBe( 2 );
		expect( sheet.getPrintSetup().getFitHeight() ).toBe( 0 );
		//xlsx
		sheet = s.getActiveSheet( xlsx );
		s.setFitToPage( xlsx, true, 2, 0 );
		expect( sheet.getFitToPage() ).toBeTrue();
		expect( sheet.getPrintSetup().getFitWidth() ).toBe( 2 );
		expect( sheet.getPrintSetup().getFitHeight() ).toBe( 0 );
	});

});	
</cfscript>