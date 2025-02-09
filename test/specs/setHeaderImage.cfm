<cfscript>
describe( "setHeaderImage", ()=>{

	it( "adds an image to the left, centre or right header from a file path", ()=>{
		var imagePath = getTestFilePath( "test.png" );
		var wb = s.newXlsx();
		s.setHeaderImage( wb, "left", imagePath );
		var header = s.getSheetHelper().getActiveSheetHeader( wb );
		expect( header.getLeft() ).toBe( "&G" );//Graphic
		wb = s.newXlsx();
		s.setHeaderImage( wb, "center", imagePath );
		header = s.getSheetHelper().getActiveSheetHeader( wb );
		expect( header.getCenter() ).toBe( "&G" );
		wb = s.newXlsx();
		s.setHeaderImage( wb, "right", imagePath );
		header = s.getSheetHelper().getActiveSheetHeader( wb );
		expect( header.getRight() ).toBe( "&G" );
	})

	it( "is chainable", ()=>{
		var imagePath = getTestFilePath( "test.png" );
		var wb = s.newXlsx();
		s.newChainable( wb ).setHeaderImage( "left", imagePath );
		var header = s.getSheetHelper().getActiveSheetHeader( wb );
		expect( header.getLeft() ).toBe( "&G" );//Graphic
	})

	it( "adds an image to the left, centre or right header from a cfml image object", ()=>{
		var imageData = ImageNew( "", 10, 10, "rgb", "blue" );
		var wb = s.newXlsx();
		s.setHeaderImage( wb, "left", imageData, "png" );
		var header = s.getSheetHelper().getActiveSheetHeader( wb );
		expect( header.getLeft() ).toBe( "&G" );//Graphic
		wb = s.newXlsx();
		s.setHeaderImage( wb, "center", imageData, "png" );
		header = s.getSheetHelper().getActiveSheetHeader( wb );
		expect( header.getCenter() ).toBe( "&G" );
		wb = s.newXlsx();
		s.setHeaderImage( wb, "right", imageData, "png" );
		header = s.getSheetHelper().getActiveSheetHeader( wb );
		expect( header.getRight() ).toBe( "&G" );
	})

	describe( "throws an exception if", ()=>{

		it( "the workbook is not XLSX", ()=>{
			expect( ()=>{
				s.setHeaderImage( s.newXls(), "left", getTestFilePath( "test.png" ) );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidSpreadsheetType" );
		})

		it( "the position argument is invalid", ()=>{
			expect( ()=>{
				s.setHeaderImage( s.newXlsx(), "wrong", getTestFilePath( "test.png" ) );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidPositionArgument" );
		})

		it( "the spreadsheet already has a header or footer image", ()=>{
			expect( ()=>{
				var wb = s.read( getTestFilePath( "hasHeaderImage.xlsx" ) );
				s.setHeaderImage( wb, "left", getTestFilePath( "test.png" ) );
			}).toThrow( type="cfsimplicity.spreadsheet.existingHeaderOrFooter" );
		})

	})	

})
</cfscript>