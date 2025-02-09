<cfscript>
describe( "deleteRows", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	})

	it( "Deletes the data in a specified range of rows", ()=>{
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ], [ "d", "e" ], [ "f", "g" ], [ "h", "i" ] ] );
		var expected = QueryNew( "column1,column2","VarChar,VarChar", [ [ "", "" ], [ "", "" ], [ "d", "e" ], [ "", "" ], [ "h", "i" ] ] );
		workbooks.Each( ( wb )=>{
			s.addRows( wb, data )
				.deleteRows( wb, "1-2,4" );
			var actual = s.getSheetHelper().sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		})
	})

	it( "Is chainable", ()=>{
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ], [ "d", "e" ], [ "f", "g" ], [ "h", "i" ] ] );
		var expected = QueryNew( "column1,column2","VarChar,VarChar", [ [ "", "" ], [ "", "" ], [ "d", "e" ], [ "", "" ], [ "h", "i" ] ] );
		workbooks.Each( ( wb )=>{
			s.newChainable( wb )
				.addRows( data )
				.deleteRows( "1-2,4" );
			var actual = s.getSheetHelper().sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		})
	})

	describe( "deleteRows throws an exception if", ()=>{

		it( "the range is invalid", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.deleteRows( wb, "a-b" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidRange" );
			})
		})

	})

})	
</cfscript>