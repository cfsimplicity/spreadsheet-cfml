<cfscript>
describe( "getCellType", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Gets the Excel data type of a cell in the active sheet", function(){
		workbooks.Each( function( wb ){
			s.setCellValue( wb, "test", 1, 1 );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
			s.setCellValue( wb, 1, 1, 1 );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
			s.setCellValue( wb,  CreateDate( 2015, 04, 12 ), 1, 1 );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
			s.setCellValue( wb, "true", 1, 1, "boolean" );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "boolean" );
			s.setCellValue( wb, "", 1, 1, "blank" );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "blank" );
			s.setCellFormula( wb, "SUM(A1:A2)", 1, 1 );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "formula" );
		});
	});

	it( "Is chainable", function(){
		workbooks.Each( function( wb ){
			var type = s.newChainable( wb )
				.setCellValue( "true", 1, 1, "boolean" )
				.getCellType( 1, 1 );
			expect( type ).toBe( "boolean" );
		});
	});

});	
</cfscript>