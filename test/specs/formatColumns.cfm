<cfscript>
describe( "formatColumns", function(){

	beforeEach( function(){
		s.clearCellStyleCache();
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			s.addRows( wb, [ [ "a1", "b1" ], [ "a2", "b2" ] ] );
		});
	});

	it(
		title="can format columns in a spreadsheet containing more than 4009 rows",
		body=function(){
			var path = getTestFilePath( "4010-rows.xls" );
			var workbook = s.read( src=path );
			var format = { italic: "true" };
			s.formatColumns( workbook, format, "1-2" );
		},
		skip=function(){
			return s.getIsACF();
		}
	);

	it( "can preserve the existing format properties other than the one(s) being changed", function(){
		workbooks.Each( function( wb ){
			s.formatColumns( wb, {  italic: true }, "1-2" );
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeTrue();
			s.formatColumns( wb, {  bold: true }, "1-2" ); //overwrites current style style by default
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeFalse();
			s.formatColumns( wb, {  italic: true }, "1-2" )
				.formatColumns( workbook=wb, format={ bold: true }, range="1-2", overwriteCurrentStyle=false );
			expect( s.getCellFormat( wb, 1, 1 ).bold ).toBeTrue();
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeTrue();
		});
	});

	it( "is chainable", function(){
		workbooks.Each( function( wb ){
			s.newChainable( wb )
				.formatColumns( { bold: true }, "1-2" );
			expect( s.getCellFormat( wb, 1, 1 ).bold ).toBeTrue();
			expect( s.getCellFormat( wb, 1, 2 ).bold ).toBeTrue();
		});
	});

	it( "works when the range is just a single column", function(){
		workbooks.Each( function( wb ){
			s.formatColumns( wb, {  italic: true }, "2" );
			expect( s.getCellFormat( wb, 2, 2 ).italic ).toBeTrue();
		});
	});

	describe( "formatColumns throws an exception if ", function(){

		it( "the range is invalid", function(){
			var workbooks = [ s.newXls(), s.newXlsx() ];
			workbooks.Each( function( wb ){
				expect( function(){
					format = { font: "Courier" };
					s.formatColumns( wb, format, "a-b" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidRange" );
			});
		});

	});

});	
</cfscript>