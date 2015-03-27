<cfscript>
describe( "setActiveSheet tests",function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Sets the specified sheet number to be active",function() {
		s.createSheet( workbook,"test" );
		makePublic( s,"getActiveSheetName" );
		s.setActiveSheet( workbook=workbook,sheetNumber=2 );
		expect( s.getActiveSheetName( workbook ) ).toBe( "test" );
	});

	it( "Sets the specified sheet name to be active",function() {
		s.createSheet( workbook,"test" );
		makePublic( s,"getActiveSheetName" );
		s.setActiveSheet( workbook=workbook,sheetName="test" );
		expect( s.getActiveSheetName( workbook ) ).toBe( "test" );
	});

	describe( "setActiveSheet exceptions",function(){

		it( "Throws an exception if the sheet name doesn't exist",function() {
			expect( function(){
				s.setActiveSheet( workbook=workbook,sheetName="test" );
			}).toThrow( regex="Invalid sheet" );
		});

		it( "Throws an exception if the sheet number doesn't exist",function() {
			expect( function(){
				s.setActiveSheet( workbook=workbook,sheetNumber=20 );
			}).toThrow( regex="Invalid sheet" );
		});

	});	

});	
</cfscript>