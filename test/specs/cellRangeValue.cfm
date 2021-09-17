<cfscript>
describe( "setCellRangeValue", function(){

	beforeEach( function(){
		variables.value = "a";
		variables.expected = querySim(
				"column1,column2
				a|a
				a|a");
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Sets the specified range of cells to the specified value", function(){
		workbooks.Each( function( wb ){
			s.setCellRangeValue( wb, value, 1, 2, 1, 2 );
			actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Is chainable", function(){
		workbooks.Each( function( wb ){
			s.newChainable( wb ).setCellRangeValue( value, 1, 2, 1, 2 );
			actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

});	
</cfscript>