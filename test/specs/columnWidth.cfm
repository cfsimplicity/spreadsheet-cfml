<cfscript>
describe( "columnWidth", function(){

	beforeEach( function(){
		var data = "a";
		variables.xls = s.new();
		variables.xlsx = s.newXlsx();
		variables.sxlsx = s.newStreamingXlsx();
		s.addRow( xls, data );
		s.addRow( xlsx, data );
		s.addRow( sxlsx, data );
	});

	it( "can set and get column width", function(){
		s.setColumnWidth( xls, 1, 10 );
		expect( s.getColumnWidth( xls, 1 ) ).toBe( 10 );
		s.setColumnWidth( xlsx, 1, 10 );
		expect( s.getColumnWidth( xlsx, 1 ) ).toBe( 10 );
		s.setColumnWidth( sxlsx, 1, 10 );
		expect( s.getColumnWidth( sxlsx, 1 ) ).toBe( 10 );
		expect( Round( s.getColumnWidthInPixels( xls, 1 ) ) ).toBe( 70 );
		expect( Round( s.getColumnWidthInPixels( xlsx, 1 ) ) ).toBe( 70 );
		expect( Round( s.getColumnWidthInPixels( sxlsx, 1 ) ) ).toBe( 70 );
	});


});	
</cfscript>