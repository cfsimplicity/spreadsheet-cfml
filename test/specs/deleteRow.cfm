<cfscript>
describe( "deleteRow", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	})

	it( "Deletes the data in a specified row", ()=>{
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "" ], [ "c", "d" ] ] );
		workbooks.Each( ( wb )=>{
			s.addRow( wb, "a,b" )
				.addRow( wb, "c,d" )
				.deleteRow( wb, 1 );
			var actual = s.getSheetHelper().sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		})
	})

	it( "Is chainable", ()=>{
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "" ], [ "c", "d" ] ] );
		workbooks.Each( ( wb )=>{
			s.newChainable( wb )
				.addRow( "a,b" )
				.addRow( "c,d" )
				.deleteRow( 1 );
			var actual = s.getSheetHelper().sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		})
	})

	describe( "deleteRow throws an exception if", ()=>{

		it( "row is zero or less", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.deleteRow( workbook=wb, row=0 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidRowArgument" );
			})
		})

	})

})	
</cfscript>