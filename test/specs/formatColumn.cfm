<cfscript>
describe( "formatColumn", function(){

	it( "can format a column containing more than 4009 rows", function(){
		var path = getTestFilePath( "4010-rows.xls" );
		var workbook = s.read( src=path );
		var format = { italic: "true" };
		s.formatColumn( workbook, format, 1 );
	});

	it( "can preserve the existing font properties when setting bold, color, font name, font size, italic, strikeout and underline", function(){
		//setup
		var xls = s.new();
		var xlsx = s.newXlsx();
		s.addColumn( xls, "a,b" );
		s.addColumn( xlsx, "a,b" );
		var format = { font: "Helvetica" };
		s.formatColumn( workbook=xls, format=format, column=1 );
		s.formatColumn( workbook=xlsx, format=format, column=1 );
		//test
		format = { bold: true };
		s.formatColumn( workbook=xls, format=format, column=1, overwriteCurrentStyle=false );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.font ).toBe( "Helvetica" );
		s.formatColumn( workbook=xlsx, format=format, column=1, overwriteCurrentStyle=false );
		cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.font ).toBe( "Helvetica" );
		//other properties already tested in formatCell
	});

	describe( "formatColumn throws an exception if", function(){

		it( "the column is 0 or below", function(){
			expect( function(){
				var workbook = s.new();
				var format = { italic="true" };
				s.formatColumn( workbook, format,0 );
			}).toThrow( regex="Invalid column" );
		});

	});

});	
</cfscript>