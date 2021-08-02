<cfscript>
describe( "createSheet", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Creates a new sheet with a unique name if name not specified", function(){
		workbooks.Each( function( wb ){
			s.createSheet( wb );
			expect( wb.getNumberOfSheets() ).toBe( 2 );
		});
	});

	it( "Creates a new sheet with the specified name", function(){
		workbooks.Each( function( wb ){
			s.createSheet( wb,"test" );
			expect( s.getSheetHelper().sheetExists( workbook=wb, sheetName="test" ) ).toBeTrue();
		});
	});

	it( "Overwrites an existing sheet with the same name if overwrite is true", function(){
		workbooks.Each( function( wb ){
			s.createSheet( wb, "test" )
				.createSheet( wb, "test", true );
			expect( wb.getNumberOfSheets() ).toBe( 2 );
		});
	});

	describe( "createSheet throws an exception if", function(){

		it( "the sheet name contains more than 31 characters", function(){
			var filename = repeatString( "a", 32 );
			workbooks.Each( function( wb ){
				expect( function(){
					s.createSheet( wb, filename );
				}).toThrow( regex="too many" );
			});
		});

		it( "the sheet name contains invalid characters", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.createSheet( wb, "[]?*\/:" );
				}).toThrow( regex="Invalid characters" );
			});
		});

		it( "a sheet exists with the specified name and overwrite is false", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.createSheet( wb, "test" )
						.createSheet( wb, "test" );
				}).toThrow( regex="already exists" );
			});
		});

	});	

});	
</cfscript>