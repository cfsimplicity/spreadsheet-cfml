<cfscript>
describe( "setActiveCell", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( ( wb )=>{
			s.addColumn( wb, "1,1" );
		})
	})

	it( "Sets the active cell on the current active sheet by default", ()=>{
		workbooks.Each( ( wb )=>{
			s.setActiveCell( wb, 2, 1 );
			expect( s.getSheetHelper().getActiveSheet( wb ).getActiveCell().toString() ).toBe( "A2" );
		})
	})

	it( "Is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			s.newChainable( wb ).setActiveCell( 2, 1 );
			expect( s.getSheetHelper().getActiveSheet( wb ).getActiveCell().toString() ).toBe( "A2" );
		})
	})

})	
</cfscript>