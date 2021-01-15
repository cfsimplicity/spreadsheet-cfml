<cfscript>
describe( "deleteColumns", function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Deletes the data in a specified range of columns", function(){
		s.addColumn( workbook, "a,b" );
		s.addColumn( workbook, "c,d" );
		s.addColumn( workbook, "e,f" );
		s.addColumn( workbook, "g,h" );
		s.addColumn( workbook, "i,j" );
		s.deleteColumns( workbook, "1-2,4" );
		var expected = querySim("column1,column2,column3,column4,column5
			||e||i
			||f||j");
		var actual = s.sheetToQuery( workbook=workbook, includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	describe( "deleteColumns throws an exception if", function(){

		it( "the range is invalid", function(){
			expect( function(){
				var workbook = s.new();
				s.deleteColumns( workbook, "a-b" );
			}).toThrow( regex="Invalid range" );
		});

	});

});	
</cfscript>