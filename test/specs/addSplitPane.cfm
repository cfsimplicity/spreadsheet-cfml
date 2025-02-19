<cfscript>
describe( "addSplitPane", ()=>{

	beforeEach( ()=>{
		var data = QueryNew( "Header1,Header2,Header3", "VarChar,VarChar,Varchar", [ [ "a", "b", "c" ], [ "d", "e", "f" ], [ "g", "h", "i" ] ] );
		var xls = s.workbookFromQuery( data );
		var xlsx = s.workbookFromQuery( data=data, xmlformat=true );
		variables.workbooks = [ xls, xlsx ];
	})

	it( "Splits a worksheet into 4 separate panes", ()=>{
		workbooks.Each( ( wb )=>{
			var sheet = s.getSheetHelper().getActiveSheet( wb );
			expect( sheet.getPaneInformation() ).toBeNull();
			s.addSplitPane( wb, 1000, 2000, 3, 2 );
			expect( sheet.getPaneInformation().isFreezePane() ).toBeFalse();
			expect( sheet.getPaneInformation().getVerticalSplitPosition() ).toBe( 1000 );
			expect( sheet.getPaneInformation().getHorizontalSplitPosition() ).toBe( 2000 );
			expect( sheet.getPaneInformation().getVerticalSplitLeftColumn() ).toBe( 3 );
			expect( sheet.getPaneInformation().getHorizontalSplitTopRow() ).toBe( 2 );
		})
	})

	/* TODO: this seems to fail with XSSF sheet.getPaneInformation().getActivePane() returns the expected byte value minus one, and a NPE if the value is 0 */
	/* it( "The active pane defaults to UPPER_LEFT", ()=>{
		workbooks.Each( ( wb )=>{
			var sheet = s.getSheetHelper().getActiveSheet( wb );
			expect( sheet.getPaneInformation() ).toBeNull();
			s.addSplitPane( wb, 1000, 2000, 1, 1 );
			expect( sheet.getPaneInformation().getActivePane() ).toBe( sheet[ "PANE_UPPER_LEFT" ] );
		})
	}) */
		
	it( "Is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			var sheet = s.getSheetHelper().getActiveSheet( wb );
			expect( sheet.getPaneInformation() ).toBeNull();
			s.newChainable( wb ).addSplitPane( 1000, 2000, 3, 2 );
			expect( sheet.getPaneInformation().isFreezePane() ).toBeFalse();
			expect( sheet.getPaneInformation().getVerticalSplitPosition() ).toBe( 1000 );
			expect( sheet.getPaneInformation().getHorizontalSplitPosition() ).toBe( 2000 );
			expect( sheet.getPaneInformation().getVerticalSplitLeftColumn() ).toBe( 3 );
			expect( sheet.getPaneInformation().getHorizontalSplitTopRow() ).toBe( 2 );
		})
	})

})	
</cfscript>