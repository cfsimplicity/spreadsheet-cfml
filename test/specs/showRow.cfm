<cfscript>
describe( "showRow",function(){

	it( "can show a row", function(){
		var query = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a","b" ], [ "c","d" ] ] );
		var xls = s.workbookFromQuery( query );
		s.hideRow( xls, 1 );
		expect( s.isRowHidden( xls, 1 ) ).toBeTrue();
		s.showRow( xls, 1 );
		expect( s.isRowHidden( xls, 1 ) ).toBeFalse();
		var xlsx = s.workbookFromQuery( data=query, xmlFormat=true );
		s.hideRow( xlsx, 1 );
		expect( s.isRowHidden( xlsx, 1 ) ).toBeTrue();
		s.showRow( xlsx, 1 );
		expect( s.isRowHidden( xlsx, 1 ) ).toBeFalse();
	});

});	
</cfscript>