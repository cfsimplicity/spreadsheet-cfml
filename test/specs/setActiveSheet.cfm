<cfscript>
describe( "setActiveSheet", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Sets the specified sheet number to be active", function(){
		workbooks.Each( function( wb ){
			s.createSheet( wb, "test" )
				.setActiveSheet( workbook=wb, sheetNumber=2 );
			expect( s.getSheetHelper().getActiveSheetName( wb ) ).toBe( "test" );
		});
	});

	it( "Sets the specified sheet name to be active", function(){
		workbooks.Each( function( wb ){
			s.createSheet( wb, "test" )
				.setActiveSheet( workbook=wb, sheetName="test" );
			expect( s.getSheetHelper().getActiveSheetName( wb ) ).toBe( "test" );
		});
	});

	it( "Is chainable", function(){
		workbooks.Each( function( wb ){
			s.newChainable( wb )
				.createSheet( "test" )
				.setActiveSheet( sheetName="test" );
			expect( s.getSheetHelper().getActiveSheetName( wb ) ).toBe( "test" );
		});
	});

	describe( "setActiveSheet throws an exception if", function(){

		it( "the sheet name doesn't exist", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.setActiveSheet( workbook=wb, sheetName="test" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetName" );
			});
		});

		it( "the sheet number doesn't exist", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.setActiveSheet( workbook=wb, sheetNumber=20 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetNumber" );
			});
		});

	});	

});	
</cfscript>