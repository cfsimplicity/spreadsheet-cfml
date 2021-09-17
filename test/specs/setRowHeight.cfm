<cfscript>
describe( "setRowHeight", function(){

	beforeEach( function(){
		var query = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
		var xls = s.workbookFromQuery( query );
		var xlsx = s.workbookFromQuery( data=query, xmlFormat=true );
		variables.workbooks = [ xls, xlsx ];
		variables.newHeight = 30;
	});

	it( "Sets the height of a row in points.", function(){
		workbooks.Each( function( wb ){
			s.setRowHeight( wb, 2, newHeight );
			var row = s.getRowHelper().getRowFromActiveSheet( wb, 2 );
			expect( row.getHeightInPoints() ).toBe( newHeight );
		});
	});

	it( "is chainable", function(){
		workbooks.Each( function( wb ){
			s.newChainable( wb ).setRowHeight( 2, newHeight );
			var row = s.getRowHelper().getRowFromActiveSheet( wb, 2 );
			expect( row.getHeightInPoints() ).toBe( newHeight );
		});
	});

});	
</cfscript>