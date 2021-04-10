<cfscript>
describe( "formatColumn", function(){

	it( "can format a column containing more than 4009 rows", function(){
		var path = getTestFilePath( "4010-rows.xls" );
		var workbook = s.read( src=path );
		var format = { italic: "true" };
		s.formatColumn( workbook, format, 1 );
	});

	it( "can preserve the existing font properties when setting bold, color, font name, font size, italic, strikeout and underline", function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			s.addColumn( wb, "a,b" );
			var format = { font: "Helvetica" };
			s.formatColumn( workbook=wb, format=format, column=1 );
			//test
			format = { bold: true };
			s.formatColumn( workbook=wb, format=format, column=1, overwriteCurrentStyle=false );
			var cellFormat = s.getCellFormat( wb, 1, 1 );
			expect( cellFormat.font ).toBe( "Helvetica" );
			//other properties already tested in formatCell
		});
	});

	describe( "formatColumn throws an exception if", function(){

		it( "the column is 0 or below", function(){
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