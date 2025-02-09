<cfscript>
describe( "deleteColumns", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	})

	it( "Deletes the data in a specified range of columns", ()=>{
		var expected = querySim("column1,column2,column3,column4,column5
			||e||i
			||f||j");
		workbooks.Each( ( wb )=>{
			s.addColumn( wb, "a,b" )
				.addColumn( wb, "c,d" )
				.addColumn( wb, "e,f" )
				.addColumn( wb, "g,h" )
				.addColumn( wb, "i,j" )
				.deleteColumns( wb, "1-2,4" );
			var actual = s.getSheetHelper().sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		})
	})

	it( "Is chainable", ()=>{
		var expected = querySim("column1,column2,column3,column4,column5
			||e||i
			||f||j");
		workbooks.Each( ( wb )=>{
			s.newChainable( wb )
				.addColumn( "a,b" )
				.addColumn( "c,d" )
				.addColumn( "e,f" )
				.addColumn( "g,h" )
				.addColumn( "i,j" )
				.deleteColumns( "1-2,4" );
			var actual = s.getSheetHelper().sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		})
	})

	describe( "deleteColumns throws an exception if", ()=>{

		it( "the range is invalid", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.deleteColumns( wb, "a-b" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidRange" );
			})
		})

	})

})	
</cfscript>