<cfscript>
describe( "columnWidth", ()=>{

	it( "can set and get column width", ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( ( wb )=>{
			s.addRow( wb, "a" )
				.setColumnWidth( wb, 1, 10 );
			expect( s.getColumnWidth( wb, 1 ) ).toBe( 10 );
			expect( Round( s.getColumnWidthInPixels( wb, 1 ) ) ).toBe( 70 );
		})
	})

	it( "getColumnWidth and setColumnWidth are chainable", ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( ( wb )=>{
			var actual = s.newChainable( wb )
				.addRow( "a" )
				.setColumnWidth( 1, 10 )
				.getColumnWidth( 1 );
			expect( actual ).toBe( 10 );
		})
	})

})	
</cfscript>