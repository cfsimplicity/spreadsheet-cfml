<cfscript>
describe( "autoSizeColumn", ()=> {

	beforeEach( ()=> {
		var data = QueryNew( "First,Last", "VarChar,VarChar", [ [ "a", "abracadabraabracadabra" ] ] );
		variables.workbook = s.workbookFromQuery( data );
	});

	it( "Doesn't error when passing valid arguments", ()=> {
		s.autoSizeColumn( workbook, 2 );
	});

	it( "Throws a helpful exception if column argument is invalid", ()=> {
		expect( ()=> {
			s.autoSizeColumn( workbook, -1 );
		}).toThrow( regex="Invalid column value" );
	});

});	
</cfscript>