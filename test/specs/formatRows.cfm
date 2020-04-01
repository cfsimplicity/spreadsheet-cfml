<cfscript>
describe( "formatRows",function(){

	it( "can preserve the existing font properties when setting bold, color, font name, font size, italic, strikeout and underline", function(){
		//setup
		xls = s.new();
		xlsx = s.newXlsx();
		s.addRows( xls,  [ [ "a", "b" ], [ "c", "d" ] ], 1, 1 );
		s.addRows( xlsx,  [ [ "a", "b" ], [ "c", "d" ] ], 1, 1 );
		var format = { font: "Helvetica" };
		s.formatRows( workbook=xls, format=format, range="1-2" );
		s.formatRows( workbook=xlsx, format=format, range="1-2" );
		//test
		format = { bold: true };
		s.formatRows( workbook=xls, format=format, range="1-2", overwriteCurrentStyle=false );
		cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.font ).toBe( "Helvetica" );
		s.formatRows( workbook=xlsx, format=format, range="1-2", overwriteCurrentStyle=false );
		cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.font ).toBe( "Helvetica" );
		//other properties already tested in formatCell
	});

	describe( "formatRows throws an exception if",function(){

		it( "the range is invalid",function() {
			expect( function(){
				workbook = s.new();
				format = { font="Courier" };
				s.formatRows( workbook,format,"a-b" );
			}).toThrow( regex="Invalid range" );
		});

	});

});	
</cfscript>