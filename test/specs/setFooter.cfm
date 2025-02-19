<cfscript>
describe( "setFooter", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	})

	it( "adds text to the left, centre or right footer", ()=>{
		workbooks.Each( ( wb )=>{
			var leftText = "I'm on the left";
			var centerText = "I'm in the middle";
			var rightText = "I'm on the right";
			s.setFooter( workbook=wb, leftFooter=leftText, centerFooter=centerText, rightFooter=rightText );
			var footer = s.getSheetHelper().getActiveSheetFooter( wb );
			expect( footer.getLeft() ).toBe( leftText );
			expect( footer.getCenter() ).toBe( centerText );
			expect( footer.getRight() ).toBe( rightText );
		})
	})

	it( "is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			var leftText = "I'm on the left";
			var centerText = "I'm in the middle";
			var rightText = "I'm on the right";
			s.newChainable( wb )
				.setFooter( leftFooter=leftText, centerFooter=centerText, rightFooter=rightText );
			var footer = s.getSheetHelper().getActiveSheetFooter( wb );
			expect( footer.getLeft() ).toBe( leftText );
			expect( footer.getCenter() ).toBe( centerText );
			expect( footer.getRight() ).toBe( rightText );
		})
	})

})	
</cfscript>