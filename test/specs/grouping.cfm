<cfscript>
describe( "column/row grouping", ()=>{

	/* Can't test visual grouping results, but check no errors */

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		var rowData = [ "a", "b", "c", "d", "e" ];
		workbooks.Each( ( wb )=>{
			s.addRows( wb, [ rowData, rowData, rowData, rowData, rowData ] );
		})
	})

	it( "can group, collapse, expand and ungroup columns", ()=>{
		workbooks.Each( ( wb )=>{
			s.groupColumns( wb, 2, 3 );
			s.collapseColumnGroup( wb, 2 );
			s.expandColumnGroup( wb, 2 );
			s.ungroupColumns( wb, 2, 3 );
		})
	})

	it( "column grouping is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			s.newChainable( wb )
				.groupColumns( 2, 3 )
				.collapseColumnGroup( 2 )
				.expandColumnGroup( 2 )
				.ungroupColumns( 2, 3 );
		})
	})

	it( "can group, collapse, expand and ungroup rows", ()=>{
		workbooks.Each( ( wb )=>{
			s.groupRows( wb, 2, 3 );
			s.collapseRowGroup( wb, 2 );
			s.expandRowGroup( wb, 2 );
			s.ungroupRows( wb, 2, 3 );
		})
	})

	it( "row grouping is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			s.newChainable( wb )
				.groupRows( 2, 3 )
				.collapseRowGroup( 2 )
				.expandRowGroup( 2 )
				.ungroupRows( 2, 3 );
		})
	})

})
</cfscript>