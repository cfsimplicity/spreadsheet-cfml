<cfscript>
describe( "hideSheet", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	})

	it( "hides the active sheet by default", ()=>{
		variables.workbooks.Each( ( wb )=>{
			s.createSheet( wb, "sheet2" ).setActiveSheet( wb, "sheet2" )
			var sheetInfo = s.sheetInfo( wb, 2 );
			expect( sheetInfo.isHidden ).toBeFalse();
			s.hideSheet( wb );
			sheetInfo = s.sheetInfo( wb, 2 );
			expect( sheetInfo.isHidden ).toBeTrue();
		})
	})

	it( "can hide a visible sheet by name or number", ()=>{
		variables.workbooks.Each( ( wb )=>{
			// by name
			s.createSheet( wb, "sheet2" );
			var sheetInfo = s.sheetInfo( wb, 2 );
			expect( sheetInfo.isHidden ).toBeFalse();
			s.hideSheet( wb, "sheet2" );
			sheetInfo = s.sheetInfo( wb, 2 );
			expect( sheetInfo.isHidden ).toBeTrue();
			// by number
			s.createSheet( wb, "sheet3" );
			var sheetInfo = s.sheetInfo( wb, 3 );
			expect( sheetInfo.isHidden ).toBeFalse();
			s.hideSheet( workbook=wb, sheetNumber=3 );
			sheetInfo = s.sheetInfo( wb, 3 );
			expect( sheetInfo.isHidden ).toBeTrue();
			// by number positionally
			s.createSheet( wb, "sheet4" );
			var sheetInfo = s.sheetInfo( wb, 4 );
			expect( sheetInfo.isHidden ).toBeFalse();
			s.hideSheet( wb, "", 4 );
			sheetInfo = s.sheetInfo( wb, 4 );
			expect( sheetInfo.isHidden ).toBeTrue();
		})
	})

	it( "prevents the hidden sheet being the active sheet", ()=>{
		variables.workbooks.Each( ( wb )=>{
			s.renameSheet( wb, "sheet1", 1 )
				.createSheet( wb, "sheet2" )
				.setActiveSheet( wb, "sheet2" )
				.hideSheet( wb, "sheet2" );
			var activeSheet = s.getSheetHelper().getActiveSheetName( wb );
			expect( activeSheet ).toBe( "sheet1" );
		})
	})

	it( "unhides the active sheet by default", ()=>{
		variables.workbooks.Each( ( wb )=>{
			s.createSheet( wb, "sheet2" )
				.hideSheet( wb, "sheet2" );
			sheetInfo = s.sheetInfo( wb, 2 );
			expect( sheetInfo.isHidden ).toBeTrue();
			s.setActiveSheet( wb, "sheet2" )
				.unhideSheet( wb );
			sheetInfo = s.sheetInfo( wb, 2 );
			expect( sheetInfo.isHidden ).toBeFalse();
		})
	})

	it( "can unhide a hidden sheet by name or number", ()=>{
		variables.workbooks.Each( ( wb )=>{
			// by name
			s.createSheet( wb, "sheet2" ).hideSheet( wb, "sheet2" );
			sheetInfo = s.sheetInfo( wb, 2 );
			expect( sheetInfo.isHidden ).toBeTrue();
			s.unhideSheet( wb, "sheet2" );
			sheetInfo = s.sheetInfo( wb, 2 );
			expect( sheetInfo.isHidden ).toBeFalse();
			// by number
			s.createSheet( wb, "sheet3" ).hideSheet( wb, "sheet3" );
			sheetInfo = s.sheetInfo( wb, 3 );
			expect( sheetInfo.isHidden ).toBeTrue();
			s.unhideSheet( workbook=wb, sheetNumber=3 );
			sheetInfo = s.sheetInfo( wb, 3 );
			expect( sheetInfo.isHidden ).toBeFalse();
			// by number positionally
			s.createSheet( wb, "sheet4" ).hideSheet( wb, "sheet4" );
			sheetInfo = s.sheetInfo( wb, 4 );
			expect( sheetInfo.isHidden ).toBeTrue();
			s.unhideSheet( wb, "", 4 );
			sheetInfo = s.sheetInfo( wb, 4 );
			expect( sheetInfo.isHidden ).toBeFalse();
		})
	})

	it( "can hide a visible sheet via chaining", ()=>{
		variables.spreadsheetTypes.Each( ( type )=>{
			var chainable = s.newChainable( type ).createSheet( "sheet2" );
			var sheetInfo = chainable.sheetInfo( 2 );
			expect( sheetInfo.isHidden ).toBeFalse();
			chainable.hideSheet( "sheet2" );
			sheetInfo = chainable.sheetInfo( 2 );
			expect( sheetInfo.isHidden ).toBeTrue();
		})
	})

	it( "can unhide a hidden sheet via chaining", ()=>{
		variables.spreadsheetTypes.Each( ( type )=>{
			var chainable = s.newChainable( type ).createSheet( "sheet2" ).hideSheet( "sheet2" );
			var sheetInfo = chainable.sheetInfo( 2 );
			expect( sheetInfo.isHidden ).toBeTrue();
			chainable.unhideSheet( "sheet2" );
			sheetInfo = chainable.sheetInfo( 2 );
			expect( sheetInfo.isHidden ).toBeFalse();
		})
	})

})	
</cfscript>