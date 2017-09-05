<cfscript>
describe( "getRowCount",function(){

	beforeEach( function(){
		variables.rowData = QueryNew( "column1,", "VarChar", [ [ "a" ], [ "c" ] ] );
	});

	it( "Can get the number of rows in the first sheet of an XLS binary workbook",function() {
		var workbook = s.newXls();
		expect( s.getRowCount( workbook ) ).toBe( 0 );// empty
		s.addRow( workbook, "A1" );
		expect( s.getRowCount( workbook ) ).toBe( 1 );
		s.addRow( workbook, "B1" );
		expect( s.getRowCount( workbook ) ).toBe( 2 );
	});

	it( "Can get the number of rows in the first sheet of an XLSX workbook",function() {
		var workbook = s.newXlsx();
		expect( s.getRowCount( workbook ) ).toBe( 0 );// empty
		s.addRow( workbook, "A1" );
		expect( s.getRowCount( workbook ) ).toBe( 1 );
		s.addRow( workbook, "B1" );
		expect( s.getRowCount( workbook ) ).toBe( 2 );
	});

	it( "Will include empty/blank rows",function() {
		var workbook = s.new();
		s.addRow( workbook, "B1", 2 );
		expect( s.getRowCount( workbook ) ).toBe( 2 );
		s.addRow( workbook, "" );
		expect( s.getRowCount( workbook ) ).toBe( 3 );
	});

	it( "Can get the number of rows of a sheet specified by number",function() {
		var workbook = s.new();
		s.createSheet( workbook );//add a second sheet and switch to it
		s.setActiveSheetNumber( workbook, 2 );
		s.addRow( workbook, "S2A1" );
		s.addRow( workbook, "S2B1" );
		s.setActiveSheetNumber( workbook,1 );//switch back to sheet 1
		s.addRow( workbook, "S1A1" );
		expect( s.getRowCount( workbook ) ).toBe( 1 );
		expect( s.getRowCount( workbook, 2 ) ).toBe( 2 );
	});

	it( "Can get the number of rows of a sheet specified by name",function() {
		var workbook = s.new();
		s.createSheet( workbook,"test" );//add a second sheet and switch to it
		s.setActiveSheetNumber( workbook, 2 );
		s.addRow( workbook, "S2A1" );
		s.addRow( workbook, "S2B1" );
		s.setActiveSheetNumber( workbook,1 );//switch back to sheet 1
		s.addRow( workbook, "S1A1" );
		expect( s.getRowCount( workbook ) ).toBe( 1 );
		expect( s.getRowCount( workbook, "test" ) ).toBe( 2 );
	});


	describe( "getRowCount throws an exception if",function(){

		it( "the sheet name or number doesn't exist",function() {
			expect( function(){
				var workbook = s.new();
				var result = s.getRowCount( workbook, 2 );
			}).toThrow( regex: "Invalid sheet" );
			expect( function(){
				var workbook = s.new();
				var result = s.getRowCount( workbook, "test" );
			}).toThrow( regex: "Invalid sheet" );
		});

	});	

});	
</cfscript>