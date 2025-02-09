<cfscript>
describe( "removeSheet", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	})

	it( "Deletes the sheet name specified", ()=>{
		workbooks.Each( ( wb )=>{
			s.createSheet( wb, "test" )
				.removeSheet( wb, "test" );
			expect( wb.getNumberOfSheets() ).toBe( 1 );
		})
	})

	it( "Is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			s.newChainable( wb )
				.createSheet( "test" )
				.removeSheet( "test" );
			expect( wb.getNumberOfSheets() ).toBe( 1 );
		})
	})

	describe( "removeSheet throws an exception if", ()=>{

		it( "the sheet name contains invalid characters", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.removeSheet( wb, "[]?*\/:" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidCharacters" );
			})
		})

		it( "the sheet name doesn't exist", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.removeSheet( wb, "test" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetName" );
			})
		})

	})	

})	
</cfscript>