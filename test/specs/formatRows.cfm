<cfscript>
describe( "formatRows", ()=>{

	beforeEach( ()=>{
		s.clearCellStyleCache();
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( ( wb )=>{
			s.addRows( wb, [ [ "a1", "b1" ], [ "a2", "b2" ] ] );
		})
	})

	it( "can preserve the existing format properties other than the one(s) being changed", ()=>{
		workbooks.Each( ( wb )=>{
			s.formatRows( wb, {  italic: true }, "1-2" );
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeTrue();
			s.formatRows( wb, {  bold: true }, "1-2" ); //overwrites current style style by default
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeFalse();
			s.formatRows( wb, {  italic: true }, "1-2" )
				.formatRows( workbook=wb, format={ bold: true }, range="1-2", overwriteCurrentStyle=false );
			expect( s.getCellFormat( wb, 1, 1 ).bold ).toBeTrue();
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeTrue();
		})
	})

	it( "is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			s.newChainable( wb )
				.formatRows( { bold: true }, "1-2" );
			expect( s.getCellFormat( wb, 1, 1 ).bold ).toBeTrue();
			expect( s.getCellFormat( wb, 2, 2 ).bold ).toBeTrue();
		})
	})

	it( "works when the range is just a single row", ()=>{
		workbooks.Each( ( wb )=>{
			s.formatRows( wb, {  italic: true }, "2" );
			expect( s.getCellFormat( wb, 2, 2 ).italic ).toBeTrue();
		})
	})

	describe( "formatRows throws an exception if", ()=>{

		it( "the range is invalid", ()=>{
			var workbooks = [ s.newXls(), s.newXlsx() ];
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					var format = { font: "Courier" };
					s.formatRows( wb, format, "a-b" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidRange" );
			})
		})

	})

})	
</cfscript>