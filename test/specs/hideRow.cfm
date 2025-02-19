<cfscript>
describe( "hideRow", ()=>{

	beforeEach( ()=>{
		var query = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
		var xls = s.workbookFromQuery( query );
		var xlsx = s.workbookFromQuery( data=query, xmlFormat=true );
		variables.workbooks = [ xls, xlsx ];
	})

	it( "can hide a row", ()=>{
		workbooks.Each( ( wb )=>{
			s.hideRow( wb, 1 );
			expect( s.isRowHidden( wb, 1 ) ).toBeTrue();
		})
	})

	it( "is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			s.newChainable( wb ).hideRow( 1 );
			expect( s.isRowHidden( wb, 1 ) ).toBeTrue();
		})
	})

})	
</cfscript>