<cfscript>
describe( "deleteRows", function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Deletes the data in a specified range of rows", function(){
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ], [ "d", "e" ], [ "f", "g" ], [ "h", "i" ] ] );
		s.addRows( workbook, data );
		s.deleteRows( workbook, "1-2,4" );
		var expected = QueryNew( "column1,column2","VarChar,VarChar", [ [ "", "" ], [ "", "" ], [ "d", "e" ], [ "", "" ], [ "h", "i" ] ] );
		var actual = s.sheetToQuery( workbook=workbook, includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	describe( "deleteRows throws an exception if", function(){

		it( "the range is invalid", function(){
			expect( function(){
				var workbook = s.new();
				s.deleteRows( workbook, "a-b" );
			}).toThrow( regex="Invalid range" );
		});

	});

});	
</cfscript>