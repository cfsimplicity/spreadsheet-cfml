<cfscript>
describe( "setSheetMargin methods", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "by default set the active sheet margins", function(){
		workbooks.Each( function( wb ){
			var sheet = s.getSheetHelper().getActiveSheet( wb );
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

	it( "are chainable", function(){
		workbooks.Each( function( wb ){
			s.newChainable( wb )
				.setSheetTopMargin( 3 )
				.setSheetBottomMargin( 3 )
				.setSheetLeftMargin( 3 )
				.setSheetRightMargin( 3 )
				.setSheetHeaderMargin( 3 )
				.setSheetFooterMargin( 3 );
			var sheet = s.getSheetHelper().getActiveSheet( wb );
			expect( sheet.getMargin( sheet.TopMargin ) ).toBe( 3 );
			expect( sheet.getMargin( sheet.BottomMargin ) ).toBe( 3 );
			expect( sheet.getMargin( sheet.LeftMargin ) ).toBe( 3 );
			expect( sheet.getMargin( sheet.RightMargin ) ).toBe( 3 );
			expect( sheet.getMargin( sheet.HeaderMargin ) ).toBe( 3 );
			expect( sheet.getMargin( sheet.FooterMargin ) ).toBe( 3 );
		});
	});

	it( "set a margin of the named sheet", function(){
		workbooks.Each( function( wb ){
			s.createSheet( wb, "test" )
				.setSheetTopMargin( wb, 3, "test" );
			var sheet = s.getSheetHelper().getSheetByName( wb, "test" );
			expect( sheet.getMargin( sheet.TopMargin ) ).toBe( 3 );
		});
	});

	it( "set a margin of the specified sheet number", function(){
		workbooks.Each( function( wb ){
			s.createSheet( wb, "test" );
			var sheet = s.getSheetHelper().getSheetByNumber( wb, 2 );
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
			var sheet = s.getSheetHelper().getActiveSheet( wb );
			s.setSheetTopMargin( wb, 3.5 );
			expect( sheet.getMargin( sheet.TopMargin ) ).toBe( 3.5 );
		});
	});

	describe( "setting sheet margins throws an exception if", function(){

		it( "both sheet name and number are specified", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.setSheetTopMargin( wb, 3, "test", 1 );
				}).toThrow( regex="Invalid arguments" );
			});
		});

	});

});	
</cfscript>