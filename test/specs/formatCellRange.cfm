<cfscript>
describe( "formatCellRange", function(){

	beforeEach( function(){
		s.clearCellStyleCache();
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			s.addRows( wb, [ [ "a1", "b1" ], [ "a2", "b2" ] ] );
		});
	});

	it( "can preserve the existing format properties other than the one(s) being changed", function(){
		workbooks.Each( function( wb ){
			s.formatCellRange( wb, { italic: true }, 1, 2, 1, 2 );
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeTrue();
			s.formatCellRange( wb, { bold: true }, 1, 2, 1, 2 ); //overwrites current style style by default
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeFalse();
			s.formatCellRange( wb, { italic: true }, 1, 2, 1, 2 )
				.formatCellRange( workbook=wb, format={ bold: true }, startRow=1, endRow=2, startColumn=1, endColumn=2, overwriteCurrentStyle=false );
			expect( s.getCellFormat( wb, 1, 1 ).bold ).toBeTrue();
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeTrue();
		});
	});

	it( "allows styles to be set using a pre-built cellStyle object", function(){
		workbooks.Each( function( wb ){
			expect( s.getCellFormat( wb, 1, 1 ).bold ).toBeFalse();
			var cellStyle = s.createCellStyle( wb, { bold: true } );
			s.formatCellRange( wb, cellStyle, 1, 2, 1, 2 );
			expect( s.getCellFormat( wb, 1, 1 ).bold ).toBeTrue();
			// support previous separate cellStyle argument without format
			cellStyle = s.createCellStyle( wb, { italic: true } );
			expect( s.getCellFormat( wb, 2, 1 ).italic ).toBeFalse();
			s.formatCellRange( workbook=wb, startRow=1, endRow=2, startColumn=1, endColumn=2, cellStyle=cellStyle );
			expect( s.getCellFormat( wb, 2, 1 ).italic ).toBeTrue();
		});
	});

	it( "Is chainable", function(){
		workbooks.Each( function( wb ){
			s.newChainable( wb )
				.formatCellRange( { bold: true }, 1, 2, 1, 2 );
			expect( s.getCellFormat( wb, 1, 1 ).bold ).toBeTrue();
			expect( s.getCellFormat( wb, 2, 2 ).bold ).toBeTrue();
		});
	});

});	
</cfscript>