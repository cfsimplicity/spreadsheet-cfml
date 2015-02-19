<cfscript>
describe( "createSheet tests",function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Creates a new sheet with a unique name if name not specified",function() {
		s.createSheet( workbook );
		expect( workbook.getNumberOfSheets() ).toBe( 2 );
	});

	it( "Creates a new sheet with the specified name",function() {
		s.createSheet( workbook,"test" );
		makePublic( s,"sheetExists" );
		expect( s.sheetExists( workbook=workbook,sheetName="test" ) ).toBeTrue();
	});

	it( "Overwrites an existing sheet with the same name if overwrite is true",function() {
		s.createSheet( workbook,"test" );
		s.createSheet( workbook,"test",true );
		expect( workbook.getNumberOfSheets() ).toBe( 2 );
	});

	describe( "createSheet exceptions",function(){

		it( "Throws an exception if the sheet name contains invalid characters",function() {
			expect( function(){
				s.createSheet( workbook,"[]?*\/:" );
			}).toThrow( regex="Invalid characters" );
		});

		it( "Throws an exception if a sheet exists with the specified name and overwrite is false",function() {
			expect( function(){
				s.createSheet( workbook,"test" );
				s.createSheet( workbook,"test" );
			}).toThrow( regex="already exists" );
		});

	});	

});	
</cfscript>