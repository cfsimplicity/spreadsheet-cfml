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

	it( "can preserve the existing format properties other than the one(s) being changed", function(){
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			s.addColumn( wb, [ "a1", "a2" ] );
		});
		workbooks.Each( function( wb ){
			s.formatColumn( wb, {  italic: true }, 1 );
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeTrue();
			s.formatColumn( wb, {  bold: true }, 1 ); //overwrites current style style by default
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeFalse();
			s.formatColumn( wb, {  italic: true }, 1 )
				.formatColumn( workbook=wb, format={ bold: true }, column=1, overwriteCurrentStyle=false );
			expect( s.getCellFormat( wb, 1, 1 ).bold ).toBeTrue();
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeTrue();
		});
	});

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