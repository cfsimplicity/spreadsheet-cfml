<cfscript>
describe( "moveSheet", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Moves the named sheet to the specified position", function(){
		workbooks.Each( function( wb ){
			s.createSheet( wb, "sheet2" );
			s.setActiveSheet( wb, "sheet2" );
			expect( s.sheetInfo( wb ).position ).toBe( 2 );
			s.moveSheet( wb, "sheet2", 1 );
			expect( s.sheetInfo( wb ).position ).toBe( 1 );
		});
	});

	it( "Is chainable", function(){
		workbooks.Each( function( wb ){
			s.newChainable( wb )
				.createSheet("sheet2" )
				.setActiveSheet( "sheet2" )
				.moveSheet( "sheet2", 1 );
			expect( s.sheetInfo( wb ).position ).toBe( 1 );
		});
	});

	describe( "moveSheet throws an exception if", function(){

		it( "the sheet name doesn't exist", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.moveSheet( wb, "test", 1 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetName" );
			});
		});

		it( "the new position is invalid", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.moveSheet( wb, "sheet1", 10 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetNumber" );
			});
		});

	});	

});	
</cfscript>