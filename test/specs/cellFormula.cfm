<cfscript>
describe( "cellFormula",function(){

	beforeEach( function(){
		variables.workbook = s.new();
		s.addColumn( workbook,"1,1" );
		variables.theFormula = "SUM(A1:A2)";
	});

	it( "Sets and gets the specified formula for the specified cell",function() {
		s.setCellFormula( workbook,theFormula, 3, 1 );
		expect( s.getCellFormula( workbook, 3, 1 ) ).toBe( theFormula );
		expect( s.getCellValue( workbook, 3, 1 ) ).toBe( 2 );
	});

	it( "Gets all formulas from the workbook",function() {
		s.setCellFormula( workbook,theFormula, 3, 1 );
		expected = [{
			formula: theFormula
			,row: 3
			,column: 1
		}];
		actual = s.getCellFormula( workbook );
		expect( actual ).toBe( expected );
	});

});	
</cfscript>