<cfscript>
describe( "setHeaderImage", function(){

	it( "adds an image to the left, centre or right header from a file path", function() {
		makePublic( s, "getActiveSheetHeader" );
		var imagePath = getTestFilePath( "test.png" );
		var wb = s.newXlsx();
		s.setHeaderImage( wb, "left", imagePath );
		var header = s.getActiveSheetHeader( wb );
		expect( header.getLeft() ).toBe( "&G" );//Graphic
		wb = s.newXlsx();
		s.setHeaderImage( wb, "center", imagePath );
		header = s.getActiveSheetHeader( wb );
		expect( header.getCenter() ).toBe( "&G" );
		wb = s.newXlsx();
		s.setHeaderImage( wb, "right", imagePath );
		header = s.getActiveSheetHeader( wb );
		expect( header.getRight() ).toBe( "&G" );
	});

	it( "adds an image to the left, centre or right header from a cfml image object", function() {
		makePublic( s, "getActiveSheetHeader" );
		var imageData = ImageNew( "", 10, 10, "rgb", "blue" );
		var wb = s.newXlsx();
		s.setHeaderImage( wb, "left", imageData, "png" );
		var header = s.getActiveSheetHeader( wb );
		expect( header.getLeft() ).toBe( "&G" );//Graphic
		wb = s.newXlsx();
		s.setHeaderImage( wb, "center", imageData, "png" );
		header = s.getActiveSheetHeader( wb );
		expect( header.getCenter() ).toBe( "&G" );
		wb = s.newXlsx();
		s.setHeaderImage( wb, "right", imageData, "png" );
		header = s.getActiveSheetHeader( wb );
		expect( header.getRight() ).toBe( "&G" );
	});

	describe( "throws an exception if", function(){

		it( "the workbook is not XLSX", function(){
			expect( function(){
				s.setHeaderImage( s.newXls(), "left", getTestFilePath( "test.png" ) );
			}).toThrow( message="Invalid spreadsheet type" );
		});

		it( "the position argument is invalid", function(){
			expect( function(){
				s.setHeaderImage( s.newXls(), "wrong", getTestFilePath( "test.png" ) );
			}).toThrow( message="Invalid header position" );
		});

	});	

});
</cfscript>