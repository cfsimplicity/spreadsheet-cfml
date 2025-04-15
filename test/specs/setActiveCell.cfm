<cfscript>
describe( "setActiveCell", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( ( wb )=>{
			s.addColumn( wb, [ 1, 1, 1 ] );
		})
	})

	it( "Sets and gets the active cell on the current active sheet by default", ()=>{
		workbooks.Each( ( wb )=>{
			s.setActiveCell( wb, 2, 1 );
			expect( s.getActiveCell( wb ) ).toBe( { column: 1, row: 2 } );
		})
	})

	it( "Is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			expect( s.newChainable( wb ).setActiveCell( 3, 1 ).getActiveCell() ).toBe( { column: 1, row: 3 } );
		})
	})

})	
</cfscript>