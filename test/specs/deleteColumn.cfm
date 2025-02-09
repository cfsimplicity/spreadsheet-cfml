<cfscript>
describe( "deleteColumn", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	})

	it( "Deletes the data in a specified column", ()=>{
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "c" ], [ "", "d" ] ] );
		workbooks.Each( ( wb )=>{
			s.addColumn( wb, "a,b" )
				.addColumn( wb, "c,d" )
				.deleteColumn( wb, 1 );
			var actual = s.getSheetHelper().sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		})
	})

	it( "Is chainable", ()=>{
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "c" ], [ "", "d" ] ] );
		workbooks.Each( ( wb )=>{
			s.newChainable( wb )
				.addColumn( "a,b" )
				.addColumn( "c,d" )
				.deleteColumn( 1 );
			var actual = s.getSheetHelper().sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		})
	})

	describe( "deleteColumn throws an exception if" , ()=>{

		it( "column is zero or less", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.deleteColumn( workbook=wb, column=0 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidColumnArgument" );
			})
		})

	})

})	
</cfscript>