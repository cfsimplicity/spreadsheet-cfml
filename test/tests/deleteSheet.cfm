<cfscript>
describe( "deleteSheet tests",function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Deletes the sheet name specified",function() {
		s.createSheet( workbook,"test" );
		s.deleteSheet( workbook=workbook,sheetName="test" );
		expect( workbook.getNumberOfSheets() ).toBe( 1 );
	});

	it( "Deletes the sheet number specified",function() {
		s.createSheet( workbook,"test" );
		s.deleteSheet( workbook=workbook,sheetNumber=2 );
		expect( workbook.getNumberOfSheets() ).toBe( 1 );
	});


	describe( "deleteSheet exceptions",function(){

		it( "Throws an exception if the sheet name contains invalid characters",function() {
			expect( function(){
				s.deleteSheet( workbook=workbook,sheetName="[]?*\/:" );
			}).toThrow( regex="Invalid characters" );
		});

		it( "Throws an exception if the sheet name doesn't exist",function() {
			expect( function(){
				s.deleteSheet( workbook=workbook,sheetName="test" );
			}).toThrow( regex="Invalid sheet" );
		});

		it( "Throws an exception if the sheet number is invalid",function() {
			expect( function(){
				s.deleteSheet( workbook=workbook,sheetNumber=0 );
			}).toThrow( regex="Invalid sheet" );
		});

		it( "Throws an exception if the sheet number doesn't exist",function() {
			expect( function(){
				s.deleteSheet( workbook=workbook,sheetNumber=20 );
			}).toThrow( regex="Invalid sheet" );
		});


	});	

});	
</cfscript>