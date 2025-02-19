<cfscript>
describe( "moveSheet", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	})

	it( "Moves the named sheet to the specified position", ()=>{
		workbooks.Each( ( wb )=>{
			s.createSheet( wb, "sheet2" );
			s.setActiveSheet( wb, "sheet2" );
			expect( s.sheetInfo( wb ).position ).toBe( 2 );
			s.moveSheet( wb, "sheet2", 1 );
			expect( s.sheetInfo( wb ).position ).toBe( 1 );
		})
	})

	it( "Is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			s.newChainable( wb )
				.createSheet("sheet2" )
				.setActiveSheet( "sheet2" )
				.moveSheet( "sheet2", 1 );
			expect( s.sheetInfo( wb ).position ).toBe( 1 );
		})
	})

	describe( "moveSheet throws an exception if", ()=>{

		it( "the sheet name doesn't exist", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.moveSheet( wb, "test", 1 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetName" );
			})
		})

		it( "the new position is invalid", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.moveSheet( wb, "sheet1", 10 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetNumber" );
			})
		})

	})	

})	
</cfscript>