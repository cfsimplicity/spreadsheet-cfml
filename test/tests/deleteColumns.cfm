<cfscript>
describe( "deleteColumns tests",function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Deletes the data in a specified range of columns",function() {
		s.addColumn( workbook,"a,b" );
		s.addColumn( workbook,"c,d" );
		s.addColumn( workbook,"e,f" );
		s.addColumn( workbook,"g,h" );
		s.addColumn( workbook,"i,j" );
		s.deleteColumns( workbook,"1-2,4" );
		expected = querySim("column1,column2,column3,column4,column5
			||e||i
			||f||j");
		actual = s.sheetToQuery( workbook=workbook,includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	describe( "deleteColumns exceptions",function(){

		it( "Throws an exception if the range is invalid",function() {
			expect( function(){
				workbook = s.new();
				s.deleteColumns( workbook,"a-b" );
			}).toThrow( regex="Invalid range" );
		});

	});

});	
</cfscript>