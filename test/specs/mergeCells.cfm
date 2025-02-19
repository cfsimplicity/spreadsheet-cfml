<cfscript>	
describe( "mergeCells", ()=>{

	beforeEach( ()=>{
		var data = querySim(
			"column1,column2
			a|b
			c|d");
		var xls = s.workbookFromQuery( data, false );
		var xlsx = s.workbookFromQuery( data=data, addHeaderRow=false, xmlFormat=true );
		variables.workbooks = [ xls, xlsx ];
	})

	it( "Retains merged cell data by default", ()=>{
		workbooks.Each( ( wb )=>{
			s.mergeCells( wb, 1, 2, 1, 2 );
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( "a" );
			expect( s.getCellValue( wb, 1, 2 ) ).toBe( "b" );
			expect( s.getCellValue( wb, 2, 1 ) ).toBe( "c" );
			expect( s.getCellValue( wb, 2, 2 ) ).toBe( "d" );
		})
	})

	it( "Can empty all but the top-left-most cell of a merged region", ()=>{
		workbooks.Each( ( wb )=>{
			s.mergeCells( wb, 1, 2, 1, 2, true );
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( "a" );
			expect( s.getCellValue( wb, 1, 2 ) ).toBe( "" );
			expect( s.getCellValue( wb, 2, 1 ) ).toBe( "" );
			expect( s.getCellValue( wb, 2, 2 ) ).toBe( "" );
		})
	})

	it( "Is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			s.newChainable( wb ).mergeCells( 1, 2, 1, 2 );
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( "a" );
			expect( s.getCellValue( wb, 1, 2 ) ).toBe( "b" );
			expect( s.getCellValue( wb, 2, 1 ) ).toBe( "c" );
			expect( s.getCellValue( wb, 2, 2 ) ).toBe( "d" );
		})
	})

	describe( "mergeCells throws an exception if", ()=>{

		it( "startRow OR startColumn is less than 1", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.mergeCells( wb, 0, 0, 1, 2 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidStartOrEndRowArgument" );
				expect( ()=>{
					s.mergeCells( wb, 1, 2, 0, 0 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidStartOrEndColumnArgument" );
			})
		})

		it( "endRow/endColumn is less than startRow/startColumn", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.mergeCells( wb, 2, 1, 1, 2 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidStartOrEndRowArgument" );
				expect( ()=>{
					s.mergeCells( wb, 1, 2, 2, 1 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidStartOrEndColumnArgument" );
			})
		})

	})
	
})	
</cfscript>