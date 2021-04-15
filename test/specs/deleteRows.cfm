<cfscript>
describe( "deleteRows", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Deletes the data in a specified range of rows", function(){
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ], [ "d", "e" ], [ "f", "g" ], [ "h", "i" ] ] );
		var expected = QueryNew( "column1,column2","VarChar,VarChar", [ [ "", "" ], [ "", "" ], [ "d", "e" ], [ "", "" ], [ "h", "i" ] ] );
		workbooks.Each( function( wb ){
			s.addRows( wb, data );
			s.deleteRows( wb, "1-2,4" );
			var actual = s.sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		});
	});

	describe( "deleteRows throws an exception if", function(){

		it( "the range is invalid", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.deleteRows( wb, "a-b" );
				}).toThrow( regex="Invalid range" );
			});
		});

	});

});	
</cfscript>