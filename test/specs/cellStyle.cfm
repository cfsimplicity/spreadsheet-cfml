<cfscript>
describe( "cellStyle", function(){

	beforeEach( function(){
		variables.xls = s.newXls();
		variables.xlsx = s.newXlsx();
		variables.format = { bold: true };
		variables.data = [ [ "x", "y" ] ];
	});

	it( "can return the total number of registered workbook cell styles", function(){
		expect( s.getWorkbookCellStylesTotal( xls ) ).toBe( 21 );
		expect( s.getWorkbookCellStylesTotal( xlsx ) ).toBe( 1 );
		s.formatColumns( xls, format, 1 );
		s.formatColumns( xlsx, format, 1 );
		expect( s.getWorkbookCellStylesTotal( xls ) ).toBe( 22 );
		expect( s.getWorkbookCellStylesTotal( xlsx ) ).toBe( 2 );
	});

	it( "can create a valid POI CellStyle object from a given format", function() {
		makePublic( s, "isValidCellStyleObject" );
		expect( s.isValidCellStyleObject( xls, s.createCellStyle( xls, format ) ) ).toBeTrue();
		expect( s.isValidCellStyleObject( xlsx, s.createCellStyle( xlsx, format ) ) ).toBeTrue();
	});

	it( "allows a single common cellStyle to be applied across multiple formatting calls and sheets", function(){
		s.addRows( xls, data );
		s.addRows( xlsx, data );
		expect( s.getWorkbookCellStylesTotal( xls ) ).toBe( 21 );
		expect( s.getWorkbookCellStylesTotal( xlsx ) ).toBe( 1 );
		xlsStyle = s.createCellStyle( xls, format );
		xlsxStyle = s.createCellStyle( xlsx, format );
		s.formatCell( workbook=xls, row=1, column=1, cellStyle=xlsStyle );
		s.formatCell( workbook=xls, row=1, column=2, cellStyle=xlsStyle );
		s.formatCell( workbook=xlsx, row=1, column=1, cellStyle=xlsxStyle );
		s.formatCell( workbook=xlsx, row=1, column=2, cellStyle=xlsxStyle );
		s.createSheet( xls );
		s.createSheet( xlsx );
		s.setActiveSheetNumber( xls, 2 );
		s.setActiveSheetNumber( xlsx, 2 );
		s.addRows( xls, data );
		s.addRows( xlsx, data );
		s.formatCell( workbook=xls, row=1, column=1, cellStyle=xlsStyle );
		s.formatCell( workbook=xls, row=1, column=2, cellStyle=xlsStyle );
		s.formatCell( workbook=xlsx, row=1, column=1, cellStyle=xlsxStyle );
		s.formatCell( workbook=xlsx, row=1, column=2, cellStyle=xlsxStyle );
		expect( s.getWorkbookCellStylesTotal( xls ) ).toBe( 22 );
		expect( s.getWorkbookCellStylesTotal( xlsx ) ).toBe( 2 );
	});

	describe( "format functions throw an exception if", function(){
		
		it( "the cellStyle argument is present but invalid", function() {
			expect( function(){
				s.formatCell( workbook=xls, row=1, column=1, cellStyle="not a cellStyle object" );
			}).toThrow( regex="Invalid argument*" );
		});

	});

});	
</cfscript>