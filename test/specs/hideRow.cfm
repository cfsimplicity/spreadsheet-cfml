<cfscript>
describe( "hideRow", function(){

	it( "can hide a row", function(){
		var query = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a","b" ], [ "c","d" ] ] );
		var xls = s.workbookFromQuery( query );
		s.hideRow( xls, 1 );
		expect( s.isRowHidden( xls, 1 ) ).toBeTrue();
		var xlsx = s.workbookFromQuery( data=query, xmlFormat=true );
		s.hideRow( xlsx, 1 );
		expect( s.isRowHidden( xlsx, 1 ) ).toBeTrue();
	});

});	
</cfscript>