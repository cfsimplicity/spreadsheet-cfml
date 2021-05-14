<cfscript>
describe( "formatColumn", function(){

	it(
		title="can format a column containing more than 4009 rows",
		body=function(){
			var path = getTestFilePath( "4010-rows.xls" );
			var workbook = s.read( src=path );
			var format = { italic: "true" };
			s.formatColumn( workbook, format, 1 );
		},
		skip=function(){
			return s.getIsACF();
		}
	);

	describe( "formatColumn throws an exception if", function(){

		it( "the column is 0 or below", function(){
			var workbooks = [ s.newXls(), s.newXlsx() ];
			workbooks.Each( function( wb ){
				expect( function(){
					var format = { italic="true" };
					s.formatColumn( wb, format,0 );
				}).toThrow( regex="Invalid column" );
			});
		});

	});

});	
</cfscript>