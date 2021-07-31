<cfscript>
describe( "cellStyle", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		variables.format = { bold: true };
		variables.data = [ [ "x", "y" ] ];
	});

	it( "can return the total number of registered workbook cell styles", function(){
		workbooks.Each( function( wb ){
			var expected = s.isXmlFormat( wb )? 1: 21;
			expect( s.getWorkbookCellStylesTotal( wb ) ).toBe( expected );
			s.formatColumns( wb, format, 1 );
			expected = s.isXmlFormat( wb )? 2: 22;
			expect( s.getWorkbookCellStylesTotal( wb ) ).toBe( expected );
		});
	});

	it( "can create a valid POI CellStyle object from a given format", function(){
		workbooks.Each( function( wb ){
			expect( s.getFormatHelper().isValidCellStyleObject( wb, s.createCellStyle( wb, format ) ) ).toBeTrue();
		});
	});

	it( "allows a single common cellStyle to be applied across multiple formatting calls and sheets", function(){
		workbooks.Each( function( wb ){
			s.addRows( wb, data );
			var expected = s.isXmlFormat( wb )? 1: 21;
			expect( s.getWorkbookCellStylesTotal( wb ) ).toBe( expected );
			var style = s.createCellStyle( wb, format );
			s.formatCell( workbook=wb, row=1, column=1, cellStyle=style );
			s.formatCell( workbook=wb, row=1, column=2, cellStyle=style );
			s.createSheet( wb );
			s.setActiveSheetNumber( wb, 2 );
			s.addRows( wb, data );
			s.formatCell( workbook=wb, row=1, column=1, cellStyle=style );
			s.formatCell( workbook=wb, row=1, column=2, cellStyle=style );
			expected = s.isXmlFormat( wb )? 2: 22;
			expect( s.getWorkbookCellStylesTotal( wb ) ).toBe( expected );
		});
	});

	describe( "format functions throw an exception if", function(){
		
		it( "the cellStyle argument is present but invalid", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.formatCell( workbook=wb, row=1, column=1, cellStyle="not a cellStyle object" );
				}).toThrow( regex="Invalid argument*" );
			});
		});

	});

});	
</cfscript>