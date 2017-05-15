<cfscript>
describe( "showColumn",function(){

	it( "can show a column", function(){
		var query = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a","b" ], [ "c","d" ] ] );
		var xls = s.workbookFromQuery( query );
		s.hideColumn( xls, 1 );
		expect( s.isColumnHidden( xls, 1 ) ).toBeTrue();
		s.showColumn( xls, 1 );
		expect( s.isColumnHidden( xls, 1 ) ).toBeFalse();
		var xlsx = s.workbookFromQuery( data=query, xmlFormat=true );
		s.hideColumn( xlsx, 1 );
		expect( s.isColumnHidden( xlsx, 1 ) ).toBeTrue();
		s.showColumn( xlsx, 1 );
		expect( s.isColumnHidden( xlsx, 1 ) ).toBeFalse();
	});

});	
</cfscript>