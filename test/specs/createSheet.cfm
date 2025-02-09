<cfscript>
describe( "createSheet", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	})

	it( "Creates a new sheet with a unique name if name not specified", ()=>{
		workbooks.Each( ( wb )=>{
			s.createSheet( wb );
			expect( wb.getNumberOfSheets() ).toBe( 2 );
		})
	})

	it( "Is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			s.newChainable( wb ).createSheet();
			expect( wb.getNumberOfSheets() ).toBe( 2 );
		})
	})

	it( "Creates a new sheet with the specified name", ()=>{
		workbooks.Each( ( wb )=>{
			s.createSheet( wb,"test" );
			expect( s.getSheetHelper().sheetExists( workbook=wb, sheetName="test" ) ).toBeTrue();
		})
	})

	it( "Overwrites an existing sheet with the same name if overwrite is true", ()=>{
		workbooks.Each( ( wb )=>{
			s.createSheet( wb, "test" )
				.createSheet( wb, "test", true );
			expect( wb.getNumberOfSheets() ).toBe( 2 );
		})
	})

	describe( "createSheet throws an exception if", ()=>{

		it( "the sheet name contains more than 31 characters", ()=>{
			var filename = repeatString( "a", 32 );
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.createSheet( wb, filename );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetName" );
			})
		})

		it( "the sheet name contains invalid characters", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.createSheet( wb, "[]?*\/:" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidCharacters" );
			})
		})

		it( "a sheet exists with the specified name and overwrite is false", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.createSheet( wb, "test" )
						.createSheet( wb, "test" );
				}).toThrow( type="cfsimplicity.spreadsheet.sheetNameAlreadyExists" );
			})
		})

	})	

})	
</cfscript>