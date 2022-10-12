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

	it( "Is chainable", function(){
		var imagePath = getTestFilePath( "test.png" );
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ) {
			s.newChainable( wb ).addImage( filepath=imagePath, anchor="1,1,2,2" );
		});
	});

	describe( "throws an exception if", function(){

		beforeEach( function(){
			variables.workbooks = [ s.newXls(), s.newXlsx() ];
		});

		it( "no image is provided", function(){
			workbooks.Each( function( wb ) {
				expect( function(){
					s.addImage( workbook=wb, anchor="1,1,2,2" );
				}).toThrow( type="cfsimplicity.spreadsheet.missingImageArgument" );
			});
		});

		it( "imageData is provided with no imageType", function(){
			workbooks.Each( function( wb ) {
				expect( function(){
					var imageData = ImageRead( getTestFilePath( "test.png" ) );
					s.addImage( workbook=wb, imageData=imageData, anchor="1,1,2,2" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidArgumentCombination" );
			});
		});

		it( "imageData is not a coldfusion image object", function(){
			workbooks.Each( function( wb ) {
				expect( function(){
					var imageData = {};
					s.addImage( workbook=wb, imageData=imageData, imageType="png", anchor="1,1,2,2" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidImage" );
			});
		});

	});	

});	
</cfscript>