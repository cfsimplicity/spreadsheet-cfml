<cfscript>
describe( "getColumnCount", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	})

	it( "Can get the maximum number of columns in the first sheet", ()=>{
		workbooks.Each( ( wb )=>{
			s.addRow( wb, "1a,1b" )
				.addRow( wb, "2a,2b,2c" )
				.addRow( wb, "3a" );
			expect( s.getColumnCount( wb ) ).toBe( 3 );
		})
	})

	it( "Is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			var count = s.newChainable( wb )
				.addRow( [ "a", "b", "c" ] )
				.getColumnCount();
			expect( count ).toBe( 3 );
		})
	})

	it( "Can get the maximum number of columns of a sheet specified by number", ()=>{
		workbooks.Each( ( wb )=>{
			s.createSheet( wb )//add a second sheet and switch to it
				.setActiveSheetNumber( wb, 2 )
				.addRow( wb, "1a,1b" )
				.addRow( wb, "2a,2b,2c" )
				.addRow( wb, "3a" )
				.setActiveSheetNumber( wb, 1 );//switch back to sheet 1
			expect( s.getColumnCount( wb ) ).toBe( 0 );
			expect( s.getColumnCount( wb, 2 ) ).toBe( 3 );
		})
	})

	it( "Can get the maximum number of columns of a sheet specified by name", ()=>{
		workbooks.Each( ( wb )=>{
			s.createSheet( wb, "test" )
				.setActiveSheetNumber( wb, 2 )
				.addRow( wb, "1a,1b" )
				.addRow( wb, "2a,2b,2c" )
				.addRow( wb, "3a" )
				.setActiveSheetNumber( wb, 1 );
			expect( s.getColumnCount( wb ) ).toBe( 0 );
			expect( s.getColumnCount( wb, "test" ) ).toBe( 3 );
		})
	})

	describe( "getColumnCount throws an exception if", ()=>{

		it( "the sheet name or number doesn't exist", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					var result=s.getColumnCount( wb, 2 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetNumber" );
			})
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					var result=s.getColumnCount( wb, "test" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetName" );
			})
		})

	})	

})	
</cfscript>