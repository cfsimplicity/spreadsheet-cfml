<cfscript>
describe( "autoSizeColumn", function(){

	beforeEach( function(){
		var data = QueryNew( "First,Last", "VarChar,VarChar", [ [ "a", "abracadabraabracadabra" ] ] );
		var xls = s.workbookFromQuery( data );
		var xlsx = s.workbookFromQuery( data=data, xmlFormat=true );
		variables.workbooks = [ xls, xlsx ];
	});

	it( "Doesn't error when passing valid arguments", function(){
		workbooks.Each( function( wb ){
			s.autoSizeColumn( wb, 2 );
		});
	});

	it( "Is chainable", function(){
		workbooks.Each( function( wb ){
			s.newChainable( wb ).autoSizeColumn( 2 );
		});
	});

	it( "Doesn't error if the workbook is SXSSF", function(){
		var data = QueryNew( "First,Last", "VarChar,VarChar", [ [ "a", "abracadabraabracadabra" ] ] );
		var workbook = s.newStreamingXlsx();
		s.addRows( local.workbook, data )
			.autoSizeColumn( local.workbook, 2 );
	});

	it( "Throws a helpful exception if column argument is invalid", function(){
		workbooks.Each( function( wb ){
			expect( function(){
				s.autoSizeColumn( wb, -1 );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidColumnArgument" );
		});
	});

});	
</cfscript>