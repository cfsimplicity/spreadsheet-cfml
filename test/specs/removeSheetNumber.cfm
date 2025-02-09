<cfscript>
describe( "removeSheetNumber", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	})

	it( "Deletes the sheet number specified", ()=>{
		workbooks.Each( ( wb )=>{
			s.createSheet( wb, "test" )
				.removeSheetNumber( wb, 2 );
			expect( wb.getNumberOfSheets() ).toBe( 1 );
		})
	})

	it( "Is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			s.newChainable( wb )
				.createSheet( "test" )
				.removeSheetNumber( 2 );
			expect( wb.getNumberOfSheets() ).toBe( 1 );
		})
	})

	describe( "removeSheetNumber throws an exception if", ()=>{

		it( "the sheet number doesn't exist", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.removeSheetNumber( wb, 20 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetNumber" );
			})
		})

	})	

})	
</cfscript>