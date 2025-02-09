<cfscript>
describe( "cellStyle", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		variables.format = { bold: true };
		variables.data = [ [ "x", "y" ] ];
	})

	it( "can create a valid POI CellStyle object from a given format", ()=>{
		workbooks.Each( ( wb )=>{
			expect( s.getFormatHelper().isValidCellStyleObject( wb, s.createCellStyle( wb, format ) ) ).toBeTrue();
		})
	})

	it( "createCellStyle is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			var style = s.newChainable( wb ).createCellStyle( format );
			expect( s.getFormatHelper().isValidCellStyleObject( wb, style ) ).toBeTrue();
		})
	})

	it( "allows a single common cellStyle to be applied across multiple formatting calls and sheets", ()=>{
		workbooks.Each( ( wb )=>{
			s.addRows( wb, data );
			var expected = s.isXmlFormat( wb )? 1: 21;
			expect( s.getWorkbookCellStylesTotal( wb ) ).toBe( expected );
			var style = s.createCellStyle( wb, format );
			s.formatCell( workbook=wb, format=style, row=1, column=1 )
				.formatCell( workbook=wb, format=style, row=1, column=2 )
				.createSheet( wb )
				.setActiveSheetNumber( wb, 2 )
				.addRows( wb, data )
				.formatCell( workbook=wb, format=style, row=1, column=1 )
				.formatCell( workbook=wb, format=style, row=1, column=2 );
			expected = s.isXmlFormat( wb )? 2: 22;
			expect( s.getWorkbookCellStylesTotal( wb ) ).toBe( expected );
		})
	})

	describe( "cellStyle utilities", ()=>{

		it( "can return the total number of registered workbook cell styles", ()=>{
			workbooks.Each( ( wb )=>{
				s.addRows( wb, data );
				var expected = s.isXmlFormat( wb )? 1: 21;
				expect( s.getWorkbookCellStylesTotal( wb ) ).toBe( expected );
				s.formatColumns( wb, format, 1 );
				expected = s.isXmlFormat( wb )? 2: 22;
				expect( s.getWorkbookCellStylesTotal( wb ) ).toBe( expected );
			})
		})

		it( "can clear the cellStyle object cache", ()=>{
			s.clearCellStyleCache();
			expect( s.getCellStyleCache().xls ).toBeEmpty();
			expect( s.getCellStyleCache().xlsx ).toBeEmpty();
			spreadsheetTypes.Each( ( type )=>{
				s.newChainable( type )
					.addRows( data )
					.formatRow( { bold: true }, 1 );
				expect( s.getCellStyleCache()[ type ] ).notToBeEmpty();
			})
			s.clearCellStyleCache();
			expect( s.getCellStyleCache().xls ).toBeEmpty();
			expect( s.getCellStyleCache().xlsx ).toBeEmpty();
		})

	})

	describe( "format functions throw an exception if", ()=>{
		
		it( "the cellStyle argument is present but invalid", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.formatCell( workbook=wb, format="not a cellStyle object", row=1, column=1 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidCellStyleArgument" );
			})
		})

	})

})	
</cfscript>