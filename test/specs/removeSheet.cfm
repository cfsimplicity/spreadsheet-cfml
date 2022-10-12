<cfscript>
describe( "removeSheet", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Deletes the sheet name specified", function(){
		workbooks.Each( function( wb ){
			s.createSheet( wb, "test" )
				.removeSheet( wb, "test" );
			expect( wb.getNumberOfSheets() ).toBe( 1 );
		});
	});

	it( "Is chainable", function(){
		workbooks.Each( function( wb ){
			s.newChainable( wb )
				.createSheet( "test" )
				.removeSheet( "test" );
			expect( wb.getNumberOfSheets() ).toBe( 1 );
		});
	});

	describe( "removeSheet throws an exception if", function(){

		it( "the sheet name contains invalid characters", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.removeSheet( wb, "[]?*\/:" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidCharacters" );
			});
		});

		it( "the sheet name doesn't exist", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.removeSheet( wb, "test" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetName" );
			});
		});

	});	

});	
</cfscript>