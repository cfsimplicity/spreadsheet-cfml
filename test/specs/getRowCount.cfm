<cfscript>
describe( "getRowCount/getLastRowNumber", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		variables.rowData = QueryNew( "column1,", "VarChar", [ [ "a" ], [ "c" ] ] );
	});

	it( "Can get the number of rows in the first sheet", function(){
		workbooks.Each( function( wb ){
			expect( s.getRowCount( wb ) ).toBe( 0 );// empty
			expect( s.getLastRowNumber( wb ) ).toBe( 0 );// empty
			s.addRow( wb, "A1" );
			expect( s.getRowCount( wb ) ).toBe( 1 );
			expect( s.getLastRowNumber( wb ) ).toBe( 1 );
			s.addRow( wb, "B1" );
			expect( s.getRowCount( wb ) ).toBe( 2 );
			expect( s.getLastRowNumber( wb ) ).toBe( 2 );
		});
	});

	it( "Are chainable", function(){
		workbooks.Each( function( wb ){
			var count = s.newChainable( wb )
				.addRow( "A1" )
				.getRowCount();
			expect( count ).toBe( 1 );
			var lastRowNum = s.newChainable( wb ).getLastRowNumber();
			expect( lastRowNum ).toBe( 1 );
		});
	});

	it( "Will include empty/blank rows", function(){
		workbooks.Each( function( wb ){
			s.addRow( wb, "B1", 2 );
			expect( s.getRowCount( wb ) ).toBe( 2 );
			expect( s.getLastRowNumber( wb ) ).toBe( 2 );
			s.addRow( wb, "" );
			expect( s.getRowCount( wb ) ).toBe( 3 );
			expect( s.getLastRowNumber( wb ) ).toBe( 3 );
		});
	});

	it( "Can get the number of rows of a sheet specified by number", function(){
		workbooks.Each( function( wb ){
			s.createSheet( wb )//add a second sheet and switch to it
				.setActiveSheetNumber( wb, 2 )
				.addRow( wb, "S2A1" )
				.addRow( wb, "S2B1" )
				.setActiveSheetNumber( wb,1 )//switch back to sheet 1
				.addRow( wb, "S1A1" );
			expect( s.getRowCount( wb ) ).toBe( 1 );
			expect( s.getLastRowNumber( wb ) ).toBe( 1 );
			expect( s.getRowCount( wb, 2 ) ).toBe( 2 );
			expect( s.getLastRowNumber( wb, 2 ) ).toBe( 2 );
		});
	});

	it( "Can get the number of rows of a sheet specified by name", function(){
		workbooks.Each( function( wb ){
			s.createSheet( wb, "test" )//add a second sheet and switch to it
				.setActiveSheetNumber( wb, 2 )
				.addRow( wb, "S2A1" )
				.addRow( wb, "S2B1" )
				.setActiveSheetNumber( wb, 1 )//switch back to sheet 1
				.addRow( wb, "S1A1" );
			expect( s.getRowCount( wb ) ).toBe( 1 );
			expect( s.getLastRowNumber( wb ) ).toBe( 1 );
			expect( s.getRowCount( wb, "test" ) ).toBe( 2 );
			expect( s.getLastRowNumber( wb, "test" ) ).toBe( 2 );
		});
	});


	describe( "getRowCount throws an exception if", function(){

		it( "the sheet name or number doesn't exist", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					var result = s.getRowCount( wb, 2 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetNumber" );
				expect( function(){
					var result = s.getLastRowNumber( wb, 2 );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetNumber" );
			});
			workbooks.Each( function( wb ){
				expect( function(){
					var result = s.getRowCount( wb, "test" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetName" );
				expect( function(){
					var result = s.getLastRowNumber( wb, "test" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetName" );
			});
		});

	});	

});	
</cfscript>