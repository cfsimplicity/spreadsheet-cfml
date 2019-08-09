<cfscript>
describe( "setSheetMargins", function(){

	beforeEach( function(){
		variables.xls = s.new();
		variables.xlsx = s.newXlsx();
	});

	it( "by default sets the active sheet margins", function() {
		makePublic( s, "getActiveSheet" );
		var sheet = s.getActiveSheet( xls );
		s.setSheetTopMargin( xls, 3 );
		expect( sheet.getMargin( sheet.TopMargin ) ).toBe( 3 );
		s.setSheetBottomMargin( xls, 3 );
		expect( sheet.getMargin( sheet.BottomMargin ) ).toBe( 3 );
		s.setSheetLeftMargin( xls, 3 );
		expect( sheet.getMargin( sheet.LeftMargin ) ).toBe( 3 );
		s.setSheetRightMargin( xls, 3 );
		expect( sheet.getMargin( sheet.RightMargin ) ).toBe( 3 );
		s.setSheetHeaderMargin( xls, 3 );
		expect( sheet.getMargin( sheet.HeaderMargin ) ).toBe( 3 );
		s.setSheetFooterMargin( xls, 3 );
		expect( sheet.getMargin( sheet.FooterMargin ) ).toBe( 3 );
		//xlsx
		sheet = s.getActiveSheet( xlsx );
		s.setSheetTopMargin( xlsx, 3 );
		expect( sheet.getMargin( sheet.TopMargin ) ).toBe( 3 );
		s.setSheetBottomMargin( xlsx, 3 );
		expect( sheet.getMargin( sheet.BottomMargin ) ).toBe( 3 );
		s.setSheetLeftMargin( xlsx, 3 );
		expect( sheet.getMargin( sheet.LeftMargin ) ).toBe( 3 );
		s.setSheetRightMargin( xlsx, 3 );
		expect( sheet.getMargin( sheet.RightMargin ) ).toBe( 3 );
		s.setSheetHeaderMargin( xlsx, 3 );
		expect( sheet.getMargin( sheet.HeaderMargin ) ).toBe( 3 );
		s.setSheetFooterMargin( xlsx, 3 );
		expect( sheet.getMargin( sheet.FooterMargin ) ).toBe( 3 );
	});

	it( "sets a margin of the named sheet", function() {
		makePublic( s, "getSheetByName" );
		s.createSheet( xls, "test" );
		s.setSheetTopMargin( xls, 3, "test" );
		var sheet = s.getSheetByName( xls, "test" );
		expect( sheet.getMargin( sheet.TopMargin ) ).toBe( 3 );
		// xlsx
		s.createSheet( xlsx, "test" );
		s.setSheetTopMargin( xlsx, 3, "test" );
		sheet = s.getSheetByName( xlsx, "test" );
		expect( sheet.getMargin( sheet.TopMargin ) ).toBe( 3 );
	});

	it( "sets a margin of the specified sheet number", function() {
		makePublic( s, "getSheetByNumber" );
		s.createSheet( xls, "test" );
		var sheet = s.getSheetByNumber( xls, 2 );
		// named arguments
		s.setSheetTopMargin( workbook=xls, marginSize=3, sheetNumber=2 );
		expect( sheet.getMargin( sheet.TopMargin ) ).toBe( 3 );
		//positional
		s.setSheetTopMargin( xls, 4, "", 2 );
		expect( sheet.getMargin( sheet.TopMargin ) ).toBe( 4 );
		//xlsx
		s.createSheet( xlsx, "test" );
		sheet = s.getSheetByNumber( xlsx, 2 );
		// named arguments
		s.setSheetTopMargin( workbook=xlsx, marginSize=3, sheetNumber=2 );
		expect( sheet.getMargin( sheet.TopMargin ) ).toBe( 3 );
		//positional
		s.setSheetTopMargin( xlsx, 4, "", 2 );
		expect( sheet.getMargin( sheet.TopMargin ) ).toBe( 4 );
	});

	it( "can set margins to floating point values", function(){
		makePublic( s, "getActiveSheet" );
		var sheet = s.getActiveSheet( xls );
		s.setSheetTopMargin( xls, 3.5 );
		expect( sheet.getMargin( sheet.TopMargin ) ).toBe( 3.5 );
	});

	describe( "setting sheet margins throws an exception if", function(){

		it( "the both sheet name and number are specified", function() {
			expect( function(){
				s.setSheetTopMargin( xls, 3, "test", 1 );
			}).toThrow( regex="Invalid arguments" );
		});

	});

});	
</cfscript>