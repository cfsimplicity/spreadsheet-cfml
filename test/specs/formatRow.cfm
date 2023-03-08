<cfscript>
describe( "formatRow", function(){

	beforeEach( function(){
		s.clearCellStyleCache();
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			s.addRow( wb, [ "a1", "b1" ] );
		});
	});

	it( "can preserve the existing format properties other than the one(s) being changed", function(){
		workbooks.Each( function( wb ){
			s.formatRow( wb, { italic: true }, 1 );
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeTrue();
			s.formatRow( wb, {  bold: true }, 1 ); //overwrites current style style by default
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeFalse();
			s.formatRow( wb, { italic: true }, 1 )
				.formatRow( workbook=wb, format={ bold: true }, row=1, overwriteCurrentStyle=false );
			expect( s.getCellFormat( wb, 1, 1 ).bold ).toBeTrue();
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeTrue();
		});
	});

	it( "is chainable", function() {
		workbooks.Each( function( wb ){
			s.newChainable( wb )
				.formatRow( { bold: true }, 1 );
			expect( s.getCellFormat( wb, 1, 1 ).bold ).toBeTrue();
			expect( s.getCellFormat( wb, 1, 2 ).bold ).toBeTrue();
		});
	});

});	
</cfscript>