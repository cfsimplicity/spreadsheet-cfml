<cfscript>
describe( "formatColumns", function(){

	it( "can format columns in a spreadsheet containing more than 4009 rows", function(){
		var path = getTestFilePath( "4010-rows.xls" );
		var workbook = s.read( src=path );
		var format = { italic: "true" };
		s.formatColumns( workbook, format, "1-2" );
	});

	describe( "formatColumns throws an exception if ", function(){

		it( "the range is invalid", function(){
			variables.workbooks = [ s.newXls(), s.newXlsx() ];
			workbooks.Each( function( wb ){
				expect( function(){
					format = { font: "Courier" };
					s.formatColumns( wb, format, "a-b" );
				}).toThrow( regex="Invalid range" );
			});
		});

	});

});	
</cfscript>