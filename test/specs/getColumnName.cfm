<cfscript>
describe( "getColumnName", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Gets the alphabetic column address reference of a given cell", function(){
		workbooks.Each( function( wb ){
			s.setCellValue( wb, "test", 1, 1 );
			expect( s.getColumnName( wb, 1 ) ).toBe( "A" );
		});
	});

	it( "Is chainable", function(){
		workbooks.Each( function( wb ){
			var result = s.newChainable( wb )
				.setCellValue( "test", 1, 1 )
				.getColumnName( 1 );
			expect( result ).toBe( "A" );
		});
	});

	describe( "getColumnName throws an exception if", function(){

		it( "the column doesn't exist", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					var result = s.getColumnName( wb, 2 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidColumn" );
			});
		});

	});	

});	
</cfscript>