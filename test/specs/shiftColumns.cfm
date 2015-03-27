<cfscript>
describe( "shiftColumns tests",function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Shifts columns right if offset is positive",function() {
		s.addColumn( workbook,"a,b" );
		s.addColumn( workbook,"c,d" );
		s.shiftColumns( workbook,1,1,1 )
		expected = querySim( "column1,column2
			|a
			|b
		");
		actual = s.sheetToQuery( workbook=workbook,includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	it( "Shifts columns left if offset is negative",function() {
		s.addColumn( workbook,"a,b" );
		s.addColumn( workbook,"c,d" );
		s.addColumn( workbook,"e,f" );
		s.shiftColumns( workbook,2,2,-1)
		expected = querySim( "column1,column2,column3
			c||e
			d||f");
		actual = s.sheetToQuery( workbook=workbook,includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

});	
</cfscript>