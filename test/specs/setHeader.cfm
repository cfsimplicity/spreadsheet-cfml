<cfscript>
describe( "setHeader", function(){

	it( "adds text to the left, centre or right header", function() {
		makePublic( s, "getActiveSheetHeader" );
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ) {
			var leftText = "I'm on the left";
			var centerText = "I'm in the middle";
			var rightText = "I'm on the right";
			s.setHeader( workbook=wb, leftHeader=leftText, centerHeader=centerText, rightHeader=rightText );
			var header = s.getActiveSheetHeader( wb );
			expect( header.getLeft() ).toBe( leftText );
			expect( header.getCenter() ).toBe( centerText );
			expect( header.getRight() ).toBe( rightText );
		});
	});

});
</cfscript>