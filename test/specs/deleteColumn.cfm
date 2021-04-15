<cfscript>
describe( "deleteColumn", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Deletes the data in a specified column", function(){
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "c" ], [ "", "d" ] ] );
		workbooks.Each( function( wb ){
			s.addColumn( wb, "a,b" );
			s.addColumn( wb, "c,d" );
			s.deleteColumn( wb, 1 );
			var actual = s.sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		});
	});

	describe( "deleteColumn throws an exception if" , function(){

		it( "column is zero or less", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.deleteColumn( workbook=wb, column=0 );
				}).toThrow( regex="Invalid column" );
			});
		});

	});

});	
</cfscript>