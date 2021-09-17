<cfscript>
describe( "setRepeatingRows", function(){

	beforeEach( function(){
		var query = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
		var xls = s.workbookFromQuery( query );
		var xlsx = s.workbookFromQuery( data=query, xmlFormat=true );
		variables.workbooks = [ xls, xlsx ];
	});

	it( "Specifies rows that should appear on every page when the current sheet is printed.", function(){
		workbooks.Each( function( wb ){
			// Make header repeat on every page
			s.setRepeatingRows( wb, "1:1" );
			var sheet = s.getSheetHelper().getActiveSheet( wb );
			expect( sheet.getRepeatingRows().formatAsString() ).toBe( "1:1" );
		});
	});

	it( "Is chainable", function(){
		workbooks.Each( function( wb ){
			s.newChainable( wb ).setRepeatingRows( "1:1" );
			var sheet = s.getSheetHelper().getActiveSheet( wb );
			expect( sheet.getRepeatingRows().formatAsString() ).toBe( "1:1" );
		});
	});

});	
</cfscript>