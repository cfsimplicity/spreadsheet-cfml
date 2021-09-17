<cfscript>
describe( "formatCellRange", function(){

	beforeEach( function(){
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