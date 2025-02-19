<cfscript>
describe( "setFooterImage", ()=>{

	it( "adds an image to the left, centre or right footer from a file path", ()=>{
		var imagePath = getTestFilePath( "test.png" );
		var wb = s.newXlsx();
		s.setFooterImage( wb, "left", imagePath );
		var footer = s.getSheetHelper().getActiveSheetFooter( wb );
		expect( footer.getLeft() ).toBe( "&G" );//Graphic
		wb = s.newXlsx();
		s.setFooterImage( wb, "center", imagePath );
		footer = s.getSheetHelper().getActiveSheetFooter( wb );
		expect( footer.getCenter() ).toBe( "&G" );
		wb = s.newXlsx();
		s.setFooterImage( wb, "right", imagePath );
		footer = s.getSheetHelper().getActiveSheetFooter( wb );
		expect( footer.getRight() ).toBe( "&G" );
	})

	it( "is chainable", ()=>{
		var imagePath = getTestFilePath( "test.png" );
		var wb = s.newXlsx();
		s.newChainable( wb ).setFooterImage( "left", imagePath );
		var footer = s.getSheetHelper().getActiveSheetFooter( wb );
		expect( footer.getLeft() ).toBe( "&G" );//Graphic
	})

	it( "adds an image to the left, centre or right footer from a cfml image object", ()=>{
		var imageData = ImageNew( "", 10, 10, "rgb", "blue" );
		var wb = s.newXlsx();
		s.setFooterImage( wb, "left", imageData, "png" );
		var footer = s.getSheetHelper().getActiveSheetFooter( wb );
		expect( footer.getLeft() ).toBe( "&G" );//Graphic
		wb = s.newXlsx();
		s.setFooterImage( wb, "center", imageData, "png" );
		footer = s.getSheetHelper().getActiveSheetFooter( wb );
		expect( footer.getCenter() ).toBe( "&G" );
		wb = s.newXlsx();
		s.setFooterImage( wb, "right", imageData, "png" );
		footer = s.getSheetHelper().getActiveSheetFooter( wb );
		expect( footer.getRight() ).toBe( "&G" );
	})

	describe( "throws an exception if", ()=>{

		it( "the workbook is not XLSX", ()=>{
			expect( ()=>{
				s.setFooterImage( s.newXls(), "left", getTestFilePath( "test.png" ) );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidSpreadsheetType" );
		})

		it( "the position argument is invalid", ()=>{
			expect( ()=>{
				s.setFooterImage( s.newXlsx(), "wrong", getTestFilePath( "test.png" ) );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidPositionArgument" );
		})

		it( "the spreadsheet already has a header or footer image", ()=>{
			expect( ()=>{
				var wb = s.read( getTestFilePath( "hasHeaderImage.xlsx" ) );
				s.setFooterImage( wb, "left", getTestFilePath( "test.png" ) );
			}).toThrow( type="cfsimplicity.spreadsheet.existingHeaderOrFooter" );
		})

	})	

})
</cfscript>