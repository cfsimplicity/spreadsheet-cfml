<cfscript>
describe( "setRepeatingColumns", function(){

	beforeEach( function(){
		var query = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
		var xls = s.workbookFromQuery( query );
		var xlsx = s.workbookFromQuery( data=query, xmlFormat=true );
		variables.workbooks = [ xls, xlsx ];
	});

	it( "Specifies columns that should appear on every page when the current sheet is printed.", function(){
		workbooks.Each( function( wb ){
			// Make column1 repeat on every page
			s.setRepeatingColumns( wb, "A:A" );
			var sheet = s.getSheetHelper().getActiveSheet( wb );
			expect( sheet.getRepeatingColumns().formatAsString() ).toBe( "A:A" );
		});
	});

	it( "Is chainable", function(){
		workbooks.Each( function( wb ){
			// Make column1 repeat on every page
			s.newChainable( wb ).setRepeatingColumns( "A:A" );
			var sheet = s.getSheetHelper().getActiveSheet( wb );
			expect( sheet.getRepeatingColumns().formatAsString() ).toBe( "A:A" );
		});
	});

});	
</cfscript>