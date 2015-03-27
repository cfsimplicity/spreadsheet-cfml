<cfscript>
describe( "cellFormula tests",function(){

	beforeEach( function(){
		variables.workbook = s.new();
		s.addColumn( workbook,"1,1" );
		variables.theFormula = "SUM(A1:A2)";
		s.setCellFormula( workbook,theFormula,3,1 );
	});

	it( "Sets the specified formula for the specified cell",function() {
		expect( s.getCellValue( workbook,3,1 ) ).toBe( 2 );
	});

	it( "Gets the formula from the specified cell",function() {
		expect( s.getCellFormula( workbook,3,1 ) ).toBe( theFormula );
	});

	it( "Gets all formulas from the workbook",function() {
		expected = [{
			formula=theFormula
			,row=3
			,column=1
		}];
		actual = s.getCellFormula( workbook );
		expect( actual ).toBe( expected );
	});

});	
</cfscript>