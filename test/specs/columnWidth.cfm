<cfscript>
describe( "columnWidth", function(){

	it( "can set and get column width", function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			s.addRow( wb, "a" );
			s.setColumnWidth( wb, 1, 10 );
			expect( s.getColumnWidth( wb, 1 ) ).toBe( 10 );
			expect( Round( s.getColumnWidthInPixels( wb, 1 ) ) ).toBe( 70 );
		});
	});

});	
</cfscript>