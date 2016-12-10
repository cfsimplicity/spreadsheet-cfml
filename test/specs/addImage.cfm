<cfscript>
describe( "addImage",function(){

	it( "Doesn't error when adding an image to a binary spreadsheet",function() {
		var imagePath = ExpandPath( "/root/test/files/test.png" );
		var workbook = s.newXls();
		s.addImage( workbook=workbook, filepath=imagePath, anchor="1,1,2,2" );
	});

	it( "Doesn't error when adding an image to an XML spreadsheet",function() {
		var imagePath = ExpandPath( "/root/test/files/test.png" );
		var workbook = s.newXlsx();
		s.addImage( workbook=workbook, filepath=imagePath, anchor="1,1,2,2" );
	});

});	
</cfscript>