<cfscript>
describe( "showColumn", ()=>{

	beforeEach( ()=>{
		var query = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a","b" ], [ "c","d" ] ] );
		var xls = s.workbookFromQuery( query );
		var xlsx = s.workbookFromQuery( data=query, xmlFormat=true );
		variables.workbooks = [ xls, xlsx ];
	})

	it( "can show a column", ()=>{
		workbooks.Each( ( wb )=>{
			s.hideColumn( wb, 1 );
			expect( s.isColumnHidden( wb, 1 ) ).toBeTrue();
			s.showColumn( wb, 1 );
			expect( s.isColumnHidden( wb, 1 ) ).toBeFalse();
		})
	})

	it( "is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			s.newChainable( wb ).hideColumn( 1 );
			expect( s.isColumnHidden( wb, 1 ) ).toBeTrue();
			s.newChainable( wb ).showColumn( 1 );
			expect( s.isColumnHidden( wb, 1 ) ).toBeFalse();
		})
	})

})	
</cfscript>