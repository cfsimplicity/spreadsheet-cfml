<cfscript>
describe( "setActiveSheetNumber tests",function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Sets the specified sheet number to be active",function() {
		s.createSheet( workbook,"test" );
		makePublic( s,"getActiveSheetName" );
		s.setActiveSheetNumber( workbook,2 );
		expect( s.getActiveSheetName( workbook ) ).toBe( "test" );
	});


	describe( "setActiveSheetNumber exceptions",function(){

		it( "Throws an exception if the sheet number doesn't exist",function() {
			expect( function(){
				s.setActiveSheetNumber( workbook,20 );
			}).toThrow( regex="Invalid sheet" );
		});

	});	

});	
</cfscript>