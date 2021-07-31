<cfscript>
describe( "setHeaderImage", function(){

	it( "adds an image to the left, centre or right header from a file path", function() {
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
	});

	it( "adds an image to the left, centre or right header from a cfml image object", function() {
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
	});

	describe( "throws an exception if", function(){

		it( "the workbook is not XLSX", function(){
			expect( function(){
				s.setHeaderImage( s.newXls(), "left", getTestFilePath( "test.png" ) );
			}).toThrow( regex="Invalid spreadsheet type" );
		});

		it( "the position argument is invalid", function(){
			expect( function(){
				s.setHeaderImage( s.newXlsx(), "wrong", getTestFilePath( "test.png" ) );
			}).toThrow( regex="Invalid header position" );
		});

		it( "the spreadsheet already has a header or footer image", function(){
			expect( function(){
				var wb = s.read( getTestFilePath( "hasHeaderImage.xlsx" ) );
				s.setHeaderImage( wb, "left", getTestFilePath( "test.png" ) );
			}).toThrow( regex="Spreadsheet contains an existing header or footer" );
		});

	});	

});
</cfscript>