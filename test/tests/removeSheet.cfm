<cfscript>
describe( "removeSheet tests",function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Deletes the sheet name specified",function() {
		s.createSheet( workbook,"test" );
		s.removeSheet( workbook=workbook,sheetName="test" );
		expect( workbook.getNumberOfSheets() ).toBe( 1 );
	});


	describe( "removeSheet exceptions",function(){

		it( "Throws an exception if the sheet name contains invalid characters",function() {
			expect( function(){
				s.removeSheet( workbook=workbook,sheetName="[]?*\/:" );
			}).toThrow( regex="Invalid characters" );
		});

		it( "Throws an exception if the sheet name doesn't exist",function() {
			expect( function(){
				s.removeSheet( workbook=workbook,sheetName="test" );
			}).toThrow( regex="Invalid sheet" );
		});

	});	

});	
</cfscript>