<cfscript>
describe( "showRow", ()=>{

	beforeEach( ()=>{
		var query = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a","b" ], [ "c","d" ] ] );
		var xls = s.workbookFromQuery( query );
		var xlsx = s.workbookFromQuery( data=query, xmlFormat=true );
		variables.workbooks = [ xls, xlsx ];
	})

	it( "can show a row", ()=>{
		workbooks.Each( function( wb ){
			s.hideRow( wb, 1 );
			expect( s.isRowHidden( wb, 1 ) ).toBeTrue();
			s.showRow( wb, 1 );
			expect( s.isRowHidden( wb, 1 ) ).toBeFalse();
		})
	})

	it( "is chainable", ()=>{
		workbooks.Each( function( wb ){
			s.newChainable( wb ).hideRow( 1 );
			expect( s.isRowHidden( wb, 1 ) ).toBeTrue();
			s.newChainable( wb ).showRow( 1 );
			expect( s.isRowHidden( wb, 1 ) ).toBeFalse();
		})
	})

})	
</cfscript>