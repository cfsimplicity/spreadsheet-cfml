<cfscript>
describe( "setFooterImage", function(){

	it( "adds an image to the left, centre or right footer from a file path", function() {
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
	});

	it( "adds an image to the left, centre or right footer from a cfml image object", function() {
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
	});

	describe( "throws an exception if", function(){

		it( "the workbook is not XLSX", function(){
			expect( function(){
				s.setFooterImage( s.newXls(), "left", getTestFilePath( "test.png" ) );
			}).toThrow( regex="Invalid spreadsheet type" );
		});

		it( "the position argument is invalid", function(){
			expect( function(){
				s.setFooterImage( s.newXlsx(), "wrong", getTestFilePath( "test.png" ) );
			}).toThrow( regex="Invalid footer position" );
		});

		it( "the spreadsheet already has a header or footer image", function(){
			expect( function(){
				var wb = s.read( getTestFilePath( "hasHeaderImage.xlsx" ) );
				s.setFooterImage( wb, "left", getTestFilePath( "test.png" ) );
			}).toThrow( regex="Spreadsheet contains an existing header or footer" );
		});

	});	

});
</cfscript>