<cfscript>
describe( "sheetInfo", ()=>{

	it( "can get info about a specific sheet within a workbook", ()=>{
		variables.spreadsheetTypes.Each( ( type )=>{
			var defaults = {
				displaysAutomaticPageBreaks: ( type == "xlsx" )//xls=false xlsx=true
				,displaysFormulas: false
				,displaysGridlines: true
				,displaysRowAndColumnHeadings: true
				,displaysZeros: true
				,hasComments: false
				,hasDataValidations: false
				,hasMergedRegions: false
				,isCurrentActiveSheet: true
				,isHidden: false
				,isRightToLeft: false
				,name: "Sheet1"
				,numberOfDataValidations: 0
				,numberOfMergedRegions: 0
				,position: 1
				,printsFitToPage: ( type == "xls" )//xlsx=false xls=true
				,printsGridlines: false
				,printsHorizontallyCentered: false
				,printsRowAndColumnHeadings: false
				,printsVerticallyCentered: false
				,recalculateFormulasOnNextOpen: false
				,visibility: "VISIBLE"
			};
			var chainable = s.newChainable( type );
			var wb = chainable.getWorkbook();
			var actual = chainable.sheetInfo();
			expect( actual ).toBe( defaults );
			var sheet = s.getSheetHelper().getActiveSheet( wb );
			// position
			expect( chainable.sheetInfo().position ).toBe( 1 );
			// displaysAutomaticPageBreaks
			sheet.setAutoBreaks( JavaCast( "boolean", true ) );
			expect( chainable.sheetInfo().displaysAutomaticPageBreaks ).toBeTrue();
			sheet.setAutoBreaks( JavaCast( "boolean", false ) );
			expect( chainable.sheetInfo().displaysAutomaticPageBreaks ).toBeFalse();
			// displaysFormulas
			sheet.setDisplayFormulas( JavaCast( "boolean", true ) );
			expect( chainable.sheetInfo().displaysFormulas ).toBeTrue();
			sheet.setDisplayFormulas( JavaCast( "boolean", false ) );
			expect( chainable.sheetInfo().displaysFormulas ).toBeFalse();
			//displaysGridlines
			sheet.setDisplayGridlines( JavaCast( "boolean", true ) );
			expect( chainable.sheetInfo().displaysGridlines ).toBeTrue();
			sheet.setDisplayGridlines( JavaCast( "boolean", false ) );
			expect( chainable.sheetInfo().displaysGridlines ).toBeFalse();
			//displaysRowAndColumnHeadings
			sheet.setDisplayRowColHeadings( JavaCast( "boolean", true ) );
			expect( chainable.sheetInfo().displaysRowAndColumnHeadings ).toBeTrue();
			sheet.setDisplayRowColHeadings( JavaCast( "boolean", false ) );
			expect( chainable.sheetInfo().displaysRowAndColumnHeadings ).toBeFalse();
			//displaysZeros
			sheet.setDisplayZeros( JavaCast( "boolean", true ) );
			expect( chainable.sheetInfo().displaysZeros ).toBeTrue();
			sheet.setDisplayZeros( JavaCast( "boolean", false ) );
			expect( chainable.sheetInfo().displaysZeros ).toBeFalse();
			//recalculateFormulasOnNextOpen
			sheet.setForceFormulaRecalculation( JavaCast( "boolean", true ) );
			expect( chainable.sheetInfo().recalculateFormulasOnNextOpen ).toBeTrue();
			sheet.setForceFormulaRecalculation( JavaCast( "boolean", false ) );
			expect( chainable.sheetInfo().recalculateFormulasOnNextOpen ).toBeFalse();
			// isRightToLeft
			sheet.setRightToLeft( JavaCast( "boolean", true ) );
			expect( chainable.sheetInfo().isRightToLeft ).toBeTrue();
			sheet.setRightToLeft( JavaCast( "boolean", false ) );
			expect( chainable.sheetInfo().isRightToLeft ).toBeFalse();
			// printsFitToPage
			chainable.setFitToPage( true );
			expect( chainable.sheetInfo().printsFitToPage ).toBeTrue();
			chainable.setFitToPage( false );
			expect( chainable.sheetInfo().printsFitToPage ).toBeFalse();
			//printsGridlines
			chainable.addPrintGridlines();
			expect( chainable.sheetInfo().printsGridlines ).toBeTrue();
			chainable.removePrintGridlines();
			expect( chainable.sheetInfo().printsGridlines ).toBeFalse();
			//printsHorizontallyCentered
			sheet.setHorizontallyCenter( JavaCast( "boolean", true ) );
			expect( chainable.sheetInfo().printsHorizontallyCentered ).toBeTrue();
			sheet.setHorizontallyCenter( JavaCast( "boolean", false ) );
			expect( chainable.sheetInfo().printsHorizontallyCentered ).toBeFalse();
			//printsVerticallyCentered
			sheet.setVerticallyCenter( JavaCast( "boolean", true ) );
			expect( chainable.sheetInfo().printsVerticallyCentered ).toBeTrue();
			sheet.setVerticallyCenter( JavaCast( "boolean", false ) );
			expect( chainable.sheetInfo().printsVerticallyCentered ).toBeFalse();
			//printsRowAndColumnHeadings
			sheet.setPrintRowAndColumnHeadings( JavaCast( "boolean", true ) );
			expect( chainable.sheetInfo().printsRowAndColumnHeadings ).toBeTrue();
			sheet.setPrintRowAndColumnHeadings( JavaCast( "boolean", false ) );
			expect( chainable.sheetInfo().printsRowAndColumnHeadings ).toBeFalse();
			//mergedRegions
			chainable.addRow( [ "a", "b" ] ).mergeCells( 1, 1, 1, 2 );
			expect( chainable.sheetInfo().hasMergedRegions ).toBeTrue();
			expect( chainable.sheetInfo().numberOfMergedRegions ).toBe( 1 );
			// hasComments
			chainable.setCellComment( { author: "me", comment: "test" }, 1, 1 );
			expect( chainable.sheetInfo().hasComments ).toBeTrue();
			chainable.createSheet( "hidden sheet" );
			//visibility etc
			s.getSheetHelper().setVisibility( wb, 2, "VERY_HIDDEN" );
			var hiddenSheetInfo = chainable.sheetInfo( 2 );
			expect( hiddenSheetInfo.visibility ).toBe( "VERY_HIDDEN" );
			expect( hiddenSheetInfo.isHidden ).toBeTrue();
			expect( hiddenSheetInfo.name ).toBe( "hidden sheet" );
			expect( hiddenSheetInfo.isCurrentActiveSheet ).toBeFalse;
		})
	})

	it( "Returns info from the current active sheet by default", ()=>{
		variables.spreadsheetTypes.Each( ( type )=>{
			var chainable = s.newChainable( type ).createSheet( "Sheet2" ).setActiveSheet( "Sheet2" );
			var actual = chainable.sheetInfo();
			expect( actual.name ).toBe( "Sheet2" );
		})
	})

})	
</cfscript>