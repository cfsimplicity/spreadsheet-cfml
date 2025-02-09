<cfscript>
describe( "setActiveSheetNumber", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	})

	it( "Sets the specified sheet number to be active", ()=>{
		workbooks.Each( ( wb )=>{
			s.createSheet( wb, "test" )
				.setActiveSheetNumber( wb, 2 );
			expect( s.getSheetHelper().getActiveSheetName( wb ) ).toBe( "test" );
		})
	})

	it( "Is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			s.newChainable( wb )
				.createSheet( "test" )
				.setActiveSheetNumber( 2 );
			expect( s.getSheetHelper().getActiveSheetName( wb ) ).toBe( "test" );
		})
	})

	describe( "setActiveSheetNumber throws an exception if", ()=>{

		it( "the sheet number doesn't exist", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.setActiveSheetNumber( wb, 20 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetNumber" );
			})
		})

	})	

})	
</cfscript>