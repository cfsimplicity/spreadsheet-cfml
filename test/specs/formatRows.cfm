<cfscript>
describe( "formatRows", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "can preserve the existing font properties when setting bold, color, font name, font size, italic, strikeout and underline", function(){
		workbooks.Each( function( wb ){
			s.addRows( wb,  [ [ "a", "b" ], [ "c", "d" ] ], 1, 1 );
			var format = { font: "Helvetica" };
			s.formatRows( workbook=wb, format=format, range="1-2" );
			//test
			var format = { bold: true };
			s.formatRows( workbook=wb, format=format, range="1-2", overwriteCurrentStyle=false );
			var cellFormat = s.getCellFormat( wb, 1, 1 );
			expect( cellFormat.font ).toBe( "Helvetica" );
			//other properties already tested in formatCell
		});
	});

	describe( "formatRows throws an exception if", function(){

		it( "the range is invalid", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					var format = { font: "Courier" };
					s.formatRows( wb, format, "a-b" );
				}).toThrow( regex="Invalid range" );
			});
		});

	});

});	
</cfscript>