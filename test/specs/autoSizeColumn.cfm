<cfscript>
describe( "autoSizeColumn", function(){

	beforeEach( function(){
		var data = QueryNew( "First,Last", "VarChar,VarChar", [ [ "a", "abracadabraabracadabra" ] ] );
		variables.workbook = s.workbookFromQuery( data );
	});

	it( "Doesn't error when passing valid arguments", function(){
		s.autoSizeColumn( workbook, 2 );
	});

	it( "Doesn't error if the workbook is SXSSF", function(){
		var data = QueryNew( "First,Last", "VarChar,VarChar", [ [ "a", "abracadabraabracadabra" ] ] );
		var workbook = s.newStreamingXlsx();
		s.addRows( local.workbook, data );
		s.autoSizeColumn( local.workbook, 2 );
	});

	it( "Throws a helpful exception if column argument is invalid", function(){
		expect( function(){
			s.autoSizeColumn( workbook, -1 );
		}).toThrow( regex="Invalid column value" );
	});

});	
</cfscript>