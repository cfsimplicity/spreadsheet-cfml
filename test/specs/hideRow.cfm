<cfscript>
describe( "hideRow", function(){

	it( "can hide a row", function(){
		var query = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a","b" ], [ "c","d" ] ] );
		var xls = s.workbookFromQuery( query );
		var xlsx = s.workbookFromQuery( data=query, xmlFormat=true );
		var workbooks = [ xls, xlsx ];
		workbooks.Each( function( wb ){
			s.hideRow( wb, 1 );
			expect( s.isRowHidden( wb, 1 ) ).toBeTrue();
		});
	});

});	
</cfscript>