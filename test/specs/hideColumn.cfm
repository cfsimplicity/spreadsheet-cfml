<cfscript>
describe( "hideColumn", function(){

	it( "can hide a column", function(){
		var query = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
		var xls = s.workbookFromQuery( query );
		var xlsx = s.workbookFromQuery( data=query, xmlFormat=true );
		var workbooks = [ xls, xlsx ];
		workbooks.Each( function( wb ){
			s.hideColumn( wb, 1 );
			expect( s.isColumnHidden( wb, 1 ) ).toBeTrue();
		});
	});

});	
</cfscript>