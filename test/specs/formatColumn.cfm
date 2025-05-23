<cfscript>
describe( "formatColumn", ()=>{

	beforeEach( ()=>{
		s.clearCellStyleCache();
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( ( wb )=>{
			s.addColumn( wb, [ "a1", "a2" ] );
		})
	})

	it(
		title="can format a column containing more than 4009 rows",
		body=()=>{
			var path = getTestFilePath( "4010-rows.xls" );
			var workbook = s.read( src=path );
			var format = { italic: "true" };
			s.formatColumn( workbook, format, 1 );
		},
		skip=s.getIsACF()
	);

	it( "can preserve the existing format properties other than the one(s) being changed", ()=>{
		workbooks.Each( ( wb )=>{
			s.formatColumn( wb, { italic: true }, 1 );
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeTrue();
			s.formatColumn( wb, { bold: true }, 1 ); //overwrites current style style by default
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeFalse();
			s.formatColumn( wb, { italic: true }, 1 )
				.formatColumn( workbook=wb, format={ bold: true }, column=1, overwriteCurrentStyle=false );
			expect( s.getCellFormat( wb, 1, 1 ).bold ).toBeTrue();
			expect( s.getCellFormat( wb, 1, 1 ).italic ).toBeTrue();
		})
	})

	it( "is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			s.newChainable( wb )
				.formatColumn( { bold: true }, 1 );
			expect( s.getCellFormat( wb, 1, 1 ).bold ).toBeTrue();
			expect( s.getCellFormat( wb, 2, 1 ).bold ).toBeTrue();
		})
	})

	describe( "formatColumn throws an exception if", ()=>{

		it( "the column is 0 or below", ()=>{
			var workbooks = [ s.newXls(), s.newXlsx() ];
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					var format = { italic="true" };
					s.formatColumn( wb, format,0 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidColumnArgument" );
			})
		})

	})

})	
</cfscript>