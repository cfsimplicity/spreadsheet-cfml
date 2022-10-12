<cfscript>
describe( "deleteRow", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Deletes the data in a specified row", function(){
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "" ], [ "c", "d" ] ] );
		workbooks.Each( function( wb ){
			s.addRow( wb, "a,b" )
				.addRow( wb, "c,d" )
				.deleteRow( wb, 1 );
			var actual = s.getSheetHelper().sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		});
	});

	it( "Is chainable", function(){
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "" ], [ "c", "d" ] ] );
		workbooks.Each( function( wb ){
			s.newChainable( wb )
				.addRow( "a,b" )
				.addRow( "c,d" )
				.deleteRow( 1 );
			var actual = s.getSheetHelper().sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		});
	});

	describe( "deleteRow throws an exception if", function(){

		it( "row is zero or less", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.deleteRow( workbook=wb, row=0 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidRowArgument" );
			});
		});

	});

});	
</cfscript>