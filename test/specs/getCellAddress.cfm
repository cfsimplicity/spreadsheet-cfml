<cfscript>
describe( "getCellAddress", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	})

	it( "Gets the alphanumeric address reference of a given cell", ()=>{
		workbooks.Each( ( wb )=>{
			s.setCellValue( wb, "test", 1, 1 );
			expect( s.getCellAddress( wb, 1, 1 ) ).toBe( "A1" );
		})
	})

	it( "Is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			var result = s.newChainable( wb )
				.setCellValue( "test", 1, 1 )
				.getCellAddress( 1, 1 );
			expect( result ).toBe( "A1" );
		})
	})

	describe( "getCellAddress throws an exception if", ()=>{

		it( "the cell doesn't exist", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					var result = s.getCellAddress( wb, 1, 1 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidCell" );
			})
		})

	})	

})	
</cfscript>