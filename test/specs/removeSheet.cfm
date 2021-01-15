<cfscript>
describe( "removeSheet", function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Deletes the sheet name specified", function(){
		s.createSheet( workbook, "test" );
		s.removeSheet( workbook, "test" );
		expect( workbook.getNumberOfSheets() ).toBe( 1 );
	});


	describe( "removeSheet throws an exception if", function(){

		it( "the sheet name contains invalid characters", function(){
			expect( function(){
				s.removeSheet( workbook, "[]?*\/:" );
			}).toThrow( regex="Invalid characters" );
		});

		it( "the sheet name doesn't exist", function(){
			expect( function(){
				s.removeSheet( workbook, "test" );
			}).toThrow( regex="Invalid sheet" );
		});

	});	

});	
</cfscript>