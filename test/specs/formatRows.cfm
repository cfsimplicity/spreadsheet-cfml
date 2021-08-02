<cfscript>
describe( "formatRows", function(){

	it( "can preserve the existing format properties other than the one(s) being changed", function(){
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			s.addRows( wb, [ [ "a1", "b1" ], [ "a2", "b2" ] ] );
		});
		workbooks.Each( function( wb ){
			s.formatRows( wb, {  italic: true }, "1-2" );
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeTrue();
			s.formatRows( wb, {  bold: true }, "1-2" ); //overwrites current style style by default
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeFalse();
			s.formatRows( wb, {  italic: true }, "1-2" )
				.formatRows( workbook=wb, format={ bold: true }, range="1-2", overwriteCurrentStyle=false );
			expect( s.getCellFormat( wb, 1, 1 ).bold ).toBeTrue();
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeTrue();
		});
	});

	describe( "formatRows throws an exception if", function(){

		it( "the range is invalid", function(){
			var workbooks = [ s.newXls(), s.newXlsx() ];
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