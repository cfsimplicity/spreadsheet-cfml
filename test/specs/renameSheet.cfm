<cfscript>
describe( "renameSheet", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	})

	it( "Renames the specified sheet", ()=>{
		workbooks.Each( ( wb )=>{
			s.renameSheet( wb, "test", 1 );
			expect( s.getSheetHelper().sheetExists( wb, "test" ) ).toBeTrue();
		})
	})

	it( "Is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			s.newChainable( wb ).renameSheet( "test", 1 );
			expect( s.getSheetHelper().sheetExists( wb, "test" ) ).toBeTrue();
		})
	})

	describe( "renameSheet throws an exception if", ()=>{

		it( "the new sheet name already exists", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.createSheet( wb, "test" )
						.createSheet( wb, "test2" )
						.renameSheet( wb, "test2", 2 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetName" );
			})
		})

	})	

})	
</cfscript>