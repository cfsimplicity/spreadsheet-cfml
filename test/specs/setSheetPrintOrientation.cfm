<cfscript>
describe( "setSheetPrintOrientation", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "by default sets the active sheet to the specified orientation", function(){
		workbooks.Each( function( wb ){
			var sheet = s.getSheetHelper().getActiveSheet( wb );
			expect( sheet.getPrintSetup().getLandscape() ).toBeFalse();
			s.setSheetPrintOrientation( wb, "landscape" );
			expect( sheet.getPrintSetup().getLandscape() ).toBeTrue();
			s.setSheetPrintOrientation( wb, "portrait" );
			expect( sheet.getPrintSetup().getLandscape() ).toBeFalse();
		});
	});

	it( "is chainable", function(){
		workbooks.Each( function( wb ){
			var sheet = s.getSheetHelper().getActiveSheet( wb );
			expect( sheet.getPrintSetup().getLandscape() ).toBeFalse();
			s.newChainable( wb ).setSheetPrintOrientation( "landscape" );
			expect( sheet.getPrintSetup().getLandscape() ).toBeTrue();
		});
	});

	it( "sets the named sheet to the specified orientation", function(){
		workbooks.Each( function( wb ){
			s.createSheet( wb, "test" )
				.setSheetPrintOrientation( wb, "landscape", "test" );
			var sheet = s.getSheetHelper().getSheetByName( wb, "test" );
			expect( sheet.getPrintSetup().getLandscape() ).toBeTrue();
		});
	});

	it( "sets the specified sheet number to the specified orientation", function(){
		workbooks.Each( function( wb ){
			s.createSheet( wb, "test" );
			var sheet = s.getSheetHelper().getSheetByNumber( wb, 2 );
			expect( sheet.getPrintSetup().getLandscape() ).toBeFalse();
			// named arguments
			s.setSheetPrintOrientation( workbook=wb, mode="landscape", sheetNumber=2 );
			expect( sheet.getPrintSetup().getLandscape() ).toBeTrue();
			//positional
			s.setSheetPrintOrientation( wb, "portrait", "", 2 );
			expect( sheet.getPrintSetup().getLandscape() ).toBeFalse();
		});
	});

	describe( "setSheetPrintOrientation throws an exception if", function(){

		it( "the mode is invalid", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.setSheetPrintOrientation( wb, "blah" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidModeArgument" );
			});
		});

		it( "both sheet name and number are specified", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.setSheetPrintOrientation( wb, "landscape", "test", 1 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidArguments" );
			});
		});

	});

});	
</cfscript>