<cfscript>
describe( "addImage", ()=>{

	it( "Doesn't error when adding an image to a spreadsheet", ()=>{
		var imagePath = getTestFilePath( "test.png" );
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( ( wb )=>{
			s.addImage( workbook=wb, filepath=imagePath, anchor="1,1,2,2" );
			var imageData = ImageNew( "", 10, 10, "rgb", "blue" );
			s.addImage( workbook=wb, imageData=imageData, imageType="png", anchor="1,2,2,3" );
			expect( wb.getAllPictures() ).toHaveLength( 2 );
		})
	})

	it( "Is chainable", ()=>{
		var imagePath = getTestFilePath( "test.png" );
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( ( wb )=>{
			s.newChainable( wb ).addImage( filepath=imagePath, anchor="1,1,2,2" );
		})
	})

	describe( "throws an exception if", ()=>{

		beforeEach( ()=>{
			variables.workbooks = [ s.newXls(), s.newXlsx() ];
		})

		it( "no image is provided", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.addImage( workbook=wb, anchor="1,1,2,2" );
				}).toThrow( type="cfsimplicity.spreadsheet.missingImageArgument" );
			})
		})

		it( "imageData is provided with no imageType", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					var imageData = ImageRead( getTestFilePath( "test.png" ).Replace( "\", "/", "ALL" ) );//boxlang won't accept "\" https://ortussolutions.atlassian.net/browse/BL-878
					s.addImage( workbook=wb, imageData=imageData, anchor="1,1,2,2" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidArgumentCombination" );
			})
		})

		it( "imageData is not a coldfusion image object", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					var imageData = {};
					s.addImage( workbook=wb, imageData=imageData, imageType="png", anchor="1,1,2,2" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidImage" );
			})
		})

	})	

})	
</cfscript>