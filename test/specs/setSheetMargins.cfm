<cfscript>
describe( "setSheetMargins", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		makePublic( s, "getActiveSheet" );
		makePublic( s, "getSheetByName" );
		makePublic( s, "getSheetByNumber" );
	});

	it( "by default sets the active sheet margins", function(){
		workbooks.Each( function( wb ){
			var sheet = s.getActiveSheet( wb );
			s.setSheetTopMargin( wb, 3 );
			expect( sheet.getMargin( sheet.TopMargin ) ).toBe( 3 );
			s.setSheetBottomMargin( wb, 3 );
			expect( sheet.getMargin( sheet.BottomMargin ) ).toBe( 3 );
			s.setSheetLeftMargin( wb, 3 );
			expect( sheet.getMargin( sheet.LeftMargin ) ).toBe( 3 );
			s.setSheetRightMargin( wb, 3 );
			expect( sheet.getMargin( sheet.RightMargin ) ).toBe( 3 );
			s.setSheetHeaderMargin( wb, 3 );
			expect( sheet.getMargin( sheet.HeaderMargin ) ).toBe( 3 );
			s.setSheetFooterMargin( wb, 3 );
			expect( sheet.getMargin( sheet.FooterMargin ) ).toBe( 3 );
		});
	});

	it( "sets a margin of the named sheet", function(){
		workbooks.Each( function( wb ){
			s.createSheet( wb, "test" );
			s.setSheetTopMargin( wb, 3, "test" );
			var sheet = s.getSheetByName( wb, "test" );
			expect( sheet.getMargin( sheet.TopMargin ) ).toBe( 3 );
		});
	});

	it( "sets a margin of the specified sheet number", function(){
		workbooks.Each( function( wb ){
			s.createSheet( wb, "test" );
			var sheet = s.getSheetByNumber( wb, 2 );
			// named arguments
			s.setSheetTopMargin( workbook=wb, marginSize=3, sheetNumber=2 );
			expect( sheet.getMargin( sheet.TopMargin ) ).toBe( 3 );
			//positional
			s.setSheetTopMargin( wb, 4, "", 2 );
			expect( sheet.getMargin( sheet.TopMargin ) ).toBe( 4 );
		});
	});

	it( "can set margins to floating point values", function(){
		workbooks.Each( function( wb ){
			var sheet = s.getActiveSheet( wb );
			s.setSheetTopMargin( wb, 3.5 );
			expect( sheet.getMargin( sheet.TopMargin ) ).toBe( 3.5 );
		});
	});

	describe( "setting sheet margins throws an exception if", function(){

		it( "the both sheet name and number are specified", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.setSheetTopMargin( wb, 3, "test", 1 );
				}).toThrow( regex="Invalid arguments" );
			});
		});

	});

});	
</cfscript>