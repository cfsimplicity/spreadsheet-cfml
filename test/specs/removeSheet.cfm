<cfscript>
describe( "removeSheet", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Deletes the sheet name specified", function(){
		workbooks.Each( function( wb ){
			s.createSheet( wb, "test" );
			s.removeSheet( wb, "test" );
			expect( wb.getNumberOfSheets() ).toBe( 1 );
		});
	});


	describe( "removeSheet throws an exception if", function(){

		it( "the sheet name contains invalid characters", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.removeSheet( wb, "[]?*\/:" );
				}).toThrow( regex="Invalid characters" );
			});
		});

		it( "the sheet name doesn't exist", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.removeSheet( wb, "test" );
				}).toThrow( regex="Invalid sheet" );
			});
		});

	});	

});	
</cfscript>