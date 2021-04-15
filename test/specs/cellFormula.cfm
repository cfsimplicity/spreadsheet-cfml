<cfscript>
describe( "cellFormula", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			s.addColumn( wb, "1,1" );
		});
		variables.theFormula = "SUM(A1:A2)";
	});

	it( "Sets and gets the specified formula for the specified cell", function(){
		workbooks.Each( function( wb ){
			s.setCellFormula( wb, theFormula, 3, 1 );
			expect( s.getCellFormula( wb, 3, 1 ) ).toBe( theFormula );
			expect( s.getCellValue( wb, 3, 1 ) ).toBe( 2 );
		});
	});

	it( "Gets all formulas from the workbook", function(){
		workbooks.Each( function( wb ){
			s.setCellFormula( wb, theFormula, 3, 1 );
			var expected = [{
				formula: theFormula
				,row: 3
				,column: 1
			}];
			var actual = s.getCellFormula( wb );
			expect( actual ).toBe( expected );
		});
	});

	describe( "getCellFormula throws an exception if", function(){

		it( "a non-existent cell is specified", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.getCellFormula( wb, 10, 10 );
				}).toThrow( regex="Non-existent cell" );
			});
		});

	});

});	
</cfscript>