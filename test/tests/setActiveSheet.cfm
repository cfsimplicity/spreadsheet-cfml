<cfscript>
describe( "setActiveSheet tests",function(){

	it( "Sets the specified sheet number to be active",function() {
		path = ExpandPath( "/root/test/files/test.xls" );// has 2 sheets
		workbook = s.read( src=path );
		makePublic( s,"getActiveSheetName" );
		s.setActiveSheet( workbook=workbook,sheetNumber=2 );
		expect( s.getActiveSheetName( workbook ) ).toBe( "sheet2" );
	});

	it( "Sets the specified sheet name to be active",function() {
		path = ExpandPath( "/root/test/files/test.xls" );// has 2 sheets
		workbook = s.read( src=path );
		makePublic( s,"getActiveSheetName" );
		s.setActiveSheet( workbook=workbook,sheetName="sheet2" );
		expect( s.getActiveSheetName( workbook ) ).toBe( "sheet2" );
	});

});	
</cfscript>