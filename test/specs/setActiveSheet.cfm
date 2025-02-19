<cfscript>
describe( "setActiveSheet", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	})

	it( "Sets the specified sheet number to be active", ()=>{
		workbooks.Each( ( wb )=>{
			s.createSheet( wb, "test" )
				.setActiveSheet( workbook=wb, sheetNumber=2 );
			expect( s.getSheetHelper().getActiveSheetName( wb ) ).toBe( "test" );
		})
	})

	it( "Sets the specified sheet name to be active", ()=>{
		workbooks.Each( ( wb )=>{
			s.createSheet( wb, "test" )
				.setActiveSheet( workbook=wb, sheetName="test" );
			expect( s.getSheetHelper().getActiveSheetName( wb ) ).toBe( "test" );
		})
	})

	it( "Is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			s.newChainable( wb )
				.createSheet( "test" )
				.setActiveSheet( sheetName="test" );
			expect( s.getSheetHelper().getActiveSheetName( wb ) ).toBe( "test" );
		})
	})

	describe( "setActiveSheet throws an exception if", ()=>{

		it( "the sheet name doesn't exist", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.setActiveSheet( workbook=wb, sheetName="test" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetName" );
			})
		})

		it( "the sheet number doesn't exist", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.setActiveSheet( workbook=wb, sheetNumber=20 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetNumber" );
			})
		})

	})	

})	
</cfscript>