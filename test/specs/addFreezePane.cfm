<cfscript>
describe( "addFreezePane", function(){

	beforeEach( function(){
		var data = QueryNew( "Header1,Header2,Header3", "VarChar,VarChar,Varchar", [ [ "a", "b", "c" ], [ "d", "e", "f" ], [ "g", "h", "i" ] ] );
		var xls = s.workbookFromQuery( data );
		var xlsx = s.workbookFromQuery( data=data, xmlformat=true );
		variables.workbooks = [ xls, xlsx ];
	});

	it( "Creates a freezepane split horizontally and/or vertically", function() {
		workbooks.Each( function( wb ){
			var sheet = s.getSheetHelper().getActiveSheet( wb );
			expect( sheet.getPaneInformation() ).toBeNull();
			s.addFreezePane( wb, 0, 1 );
			expect( sheet.getPaneInformation().isFreezePane() ).toBeTrue();
			expect( sheet.getPaneInformation().getVerticalSplitPosition() ).toBe( 0 );
			expect( sheet.getPaneInformation().getHorizontalSplitPosition() ).toBe( 1 );
			s.addFreezePane( wb, 1, 1 );
			expect( sheet.getPaneInformation().isFreezePane() ).toBeTrue();
			expect( sheet.getPaneInformation().getVerticalSplitPosition() ).toBe( 1 );
			expect( sheet.getPaneInformation().getHorizontalSplitPosition() ).toBe( 1 );
			s.addFreezePane( wb, 1, 0 );
			expect( sheet.getPaneInformation().isFreezePane() ).toBeTrue();
			expect( sheet.getPaneInformation().getVerticalSplitPosition() ).toBe( 1 );
			expect( sheet.getPaneInformation().getHorizontalSplitPosition() ).toBe( 0 );
		});
		
	});

	it( "Can optionally set the visible left column in the right pane and/or top row in the bottom pane", function() {
		workbooks.Each( function( wb ){
			var sheet = s.getSheetHelper().getActiveSheet( wb );
			s.addFreezePane( wb, 1, 1, 3 );
			expect( sheet.getPaneInformation().isFreezePane() ).toBeTrue();
			expect( sheet.getPaneInformation().getVerticalSplitLeftColumn() ).toBe( 3 );
			expect( sheet.getPaneInformation().getHorizontalSplitTopRow() ).toBe( 1 );
			s.addFreezePane( workbook=wb, freezeColumn=1, freezeRow=1, topRow=3 );
			expect( sheet.getPaneInformation().isFreezePane() ).toBeTrue();
			expect( sheet.getPaneInformation().getVerticalSplitLeftColumn() ).toBe( 1 );
			expect( sheet.getPaneInformation().getHorizontalSplitTopRow() ).toBe( 3 );
			s.addFreezePane( wb, 1, 1, 3, 3 );
			expect( sheet.getPaneInformation().isFreezePane() ).toBeTrue();
			expect( sheet.getPaneInformation().getVerticalSplitLeftColumn() ).toBe( 3 );
			expect( sheet.getPaneInformation().getHorizontalSplitTopRow() ).toBe( 3 );
		});
		
	});

	it( "Can remove a freezepane by passing in zeros", function() {
		workbooks.Each( function( wb ){
			var sheet = s.getSheetHelper().getActiveSheet( wb );
			s.addFreezePane( wb, 0, 1 );
			expect( sheet.getPaneInformation().isFreezePane() ).toBeTrue();
			expect( sheet.getPaneInformation().getVerticalSplitPosition() ).toBe( 0 );
			expect( sheet.getPaneInformation().getHorizontalSplitPosition() ).toBe( 1 );
			s.addFreezePane( wb, 0, 0 );
			expect( sheet.getPaneInformation() ).toBeNull();
		});
		
	});

	it( "Is chainable", function() {
		workbooks.Each( function( wb ){
			var sheet = s.getSheetHelper().getActiveSheet( wb );
			expect( sheet.getPaneInformation() ).toBeNull();
			s.newChainable( wb ).addFreezePane( 0, 1 );
			expect( sheet.getPaneInformation().isFreezePane() ).toBeTrue();
			expect( sheet.getPaneInformation().getVerticalSplitPosition() ).toBe( 0 );
			expect( sheet.getPaneInformation().getHorizontalSplitPosition() ).toBe( 1 );
		});
	});

});	
</cfscript>