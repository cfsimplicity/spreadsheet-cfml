<cfscript>
describe( "sheetPrintOrientation", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	})

	it( "by default sets the active sheet to the specified orientation", ()=>{
		workbooks.Each( ( wb )=>{
			var sheet = s.getSheetHelper().getActiveSheet( wb );
			expect( s.getSheetPrintOrientation( wb ) ).toBe( "portrait" );
			s.setSheetPrintOrientation( wb, "landscape" );
			expect( s.getSheetPrintOrientation( wb ) ).toBe( "landscape" );
			s.setSheetPrintOrientation( wb, "portrait" );
			expect( s.getSheetPrintOrientation( wb ) ).toBe( "portrait" );
		})
	})

	it( "is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			var chainable = s.newChainable( wb );
			expect( chainable.getSheetPrintOrientation() ).toBe( "portrait" );
			chainable.setSheetPrintOrientation( "landscape" );
			expect( chainable.getSheetPrintOrientation() ).toBe( "landscape" );
		})
	})

	it( "sets the named sheet to the specified orientation", ()=>{
		workbooks.Each( ( wb )=>{
			s.createSheet( wb, "test" ).setSheetPrintOrientation( wb, "landscape", "test" );
			expect( s.getSheetPrintOrientation( wb, "test" ) ).toBe( "landscape" );
		})
	})

	it( "sets the specified sheet number to the specified orientation", ()=>{
		workbooks.Each( ( wb )=>{
			s.createSheet( wb, "test" );
			expect( s.getSheetPrintOrientation( workbook=wb, sheetNumber=2 ) ).toBe( "portrait" );
			// named arguments
			s.setSheetPrintOrientation( workbook=wb, mode="landscape", sheetNumber=2 );
			expect( s.getSheetPrintOrientation( workbook=wb, sheetNumber=2 ) ).toBe( "landscape" );
			//positional
			s.setSheetPrintOrientation( wb, "portrait", "", 2 );
			expect( s.getSheetPrintOrientation( wb, "", 2 ) ).toBe( "portrait" );
		})
	})

	describe( "setSheetPrintOrientation throws an exception if", ()=>{

		it( "the mode is invalid", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.setSheetPrintOrientation( wb, "blah" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidModeArgument" );
			})
		})

		it( "both sheet name and number are specified", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.setSheetPrintOrientation( wb, "landscape", "test", 1 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidArguments" );
				expect( ()=>{
					s.getSheetPrintOrientation( wb, "test", 1 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidArguments" );
			})
		})

	})

})	
</cfscript>