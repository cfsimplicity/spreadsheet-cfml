<cfscript>
describe( "getCellAddress", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Gets the alphanumeric address reference of a given cell", function(){
		workbooks.Each( function( wb ){
			s.setCellValue( wb, "test", 1, 1 );
			expect( s.getCellAddress( wb, 1, 1 ) ).toBe( "A1" );
		});
	});

	it( "Is chainable", function(){
		workbooks.Each( function( wb ){
			var result = s.newChainable( wb )
				.setCellValue( "test", 1, 1 )
				.getCellAddress( 1, 1 );
			expect( result ).toBe( "A1" );
		});
	});

	describe( "getCellAddress throws an exception if", function(){

		it( "the cell doesn't exist", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					var result = s.getCellAddress( wb, 1, 1 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidCell" );
			});
		});

	});	

});	
</cfscript>