<cfscript>
describe( "setHeader", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	})

	it( "adds text to the left, centre or right header", ()=>{
		workbooks.Each( ( wb )=>{
			var leftText = "I'm on the left";
			var centerText = "I'm in the middle";
			var rightText = "I'm on the right";
			s.setHeader( workbook=wb, leftHeader=leftText, centerHeader=centerText, rightHeader=rightText );
			var header = s.getSheetHelper().getActiveSheetHeader( wb );
			expect( header.getLeft() ).toBe( leftText );
			expect( header.getCenter() ).toBe( centerText );
			expect( header.getRight() ).toBe( rightText );
		})
	})

	it( "is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			var leftText = "I'm on the left";
			var centerText = "I'm in the middle";
			var rightText = "I'm on the right";
			s.newChainable( wb ).
				setHeader( leftHeader=leftText, centerHeader=centerText, rightHeader=rightText );
			var header = s.getSheetHelper().getActiveSheetHeader( wb );
			expect( header.getLeft() ).toBe( leftText );
			expect( header.getCenter() ).toBe( centerText );
			expect( header.getRight() ).toBe( rightText );
		})
	})

})
</cfscript>