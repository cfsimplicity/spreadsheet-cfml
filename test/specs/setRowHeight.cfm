<cfscript>
describe( "setRowHeight", ()=>{

	beforeEach( ()=>{
		var query = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
		var xls = s.workbookFromQuery( query );
		var xlsx = s.workbookFromQuery( data=query, xmlFormat=true );
		variables.workbooks = [ xls, xlsx ];
		variables.newHeight = 30;
	})

	it( "Sets the height of a row in points.", ()=>{
		workbooks.Each( ( wb )=>{
			s.setRowHeight( wb, 2, newHeight );
			var row = s.getRowHelper().getRowFromActiveSheet( wb, 2 );
			expect( row.getHeightInPoints() ).toBe( newHeight );
		})
	})

	it( "is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			s.newChainable( wb ).setRowHeight( 2, newHeight );
			var row = s.getRowHelper().getRowFromActiveSheet( wb, 2 );
			expect( row.getHeightInPoints() ).toBe( newHeight );
		})
	})

	describe( "setRowHeight throws an exception if", ()=>{

		it( "the specified row doesn't exist", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.setRowHeight( wb, 10, newHeight );
				}).toThrow( type="cfsimplicity.spreadsheet.nonExistentRow" );
			})
		})

	})

})	
</cfscript>