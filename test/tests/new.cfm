<cfscript>
describe( "new tests",function(){

	it( "Returns an HSSF workbook by default",function() {
		var workbook = s.new();
		expect( workbook.getClass().name ).toBe( "org.apache.poi.hssf.usermodel.HSSFWorkbook" );
	});

	it( "Returns an XSSF workbook if xmlFormat is true",function() {
		var workbook = s.new( xmlFormat=true );
		expect( workbook.getClass().name ).toBe( "org.apache.poi.xssf.usermodel.XSSFWorkbook" );
	});

	it( "Creates a workbook with the specified sheet name",function() {
		var workbook = s.new( "test" );
		makePublic( s,"getActiveSheetName" );
		expect( s.getActiveSheetName( workbook ) ).toBe( "test" );
	});

	describe( "new exceptions",function(){

		it( "Throws an exception if the sheet name contains invalid characters",function() {
			expect( function(){
				s.new( "[]?*\/:" );
			}).toThrow( regex="Invalid characters" );
		});

	});

});	
</cfscript>