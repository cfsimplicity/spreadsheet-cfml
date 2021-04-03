<cfscript>
describe( "addImage", function(){

	it( "Doesn't error when adding an image to a spreadsheet", function(){
		var imagePath = getTestFilePath( "test.png" );
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ) {
			s.addImage( workbook=wb, filepath=imagePath, anchor="1,1,2,2" );
			var imageData = ImageNew( "", 10, 10, "rgb", "blue" );
			s.addImage( workbook=wb, imageData=imageData, imageType="png", anchor="1,2,2,3" );
		});
	});

	describe( "throws an exception if", function(){

		beforeEach( function(){
			variables.workbook = s.newXls();
		});

		it( "no image is provided", function(){
			expect( function(){
				s.addImage( workbook=workbook, anchor="1,1,2,2" );
			}).toThrow( message="Invalid argument combination" );
		});

		it( "imageData is provided with no imageType", function(){
			expect( function(){
				var imageData = ImageRead( getTestFilePath( "test.png" ) );
				s.addImage( workbook=workbook, imageData=imageData, anchor="1,1,2,2" );
			}).toThrow( message="Invalid argument combination" );
		});

		it( "imageData is not a coldfusion image object", function(){
			expect( function(){
				var imageData = "I'm not an image";
				s.addImage( workbook=workbook, imageData=imageData, imageType="png", anchor="1,1,2,2" );
			}).toThrow( message="Invalid imageData" );
		});

	});	

});	
</cfscript>