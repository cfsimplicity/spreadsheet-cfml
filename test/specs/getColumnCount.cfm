<cfscript>
describe( "getColumnCount",function(){

	it( "Can get the maximum number of columns in the first sheet of an XLS binary workbook",function() {
		var workbook=s.new();
		s.addRow( workbook,"1a,1b" );
		s.addRow( workbook,"2a,2b,2c" );
		s.addRow( workbook,"3a" );
		expect( s.getColumnCount( workbook ) ).toBe( 3 );
	});

	it( "Can get the maximum number of columns in the first sheet of an XLSX workbook",function() {
		var workbook=s.newXlsx();
		s.addRow( workbook,"1a,1b" );
		s.addRow( workbook,"2a,2b,2c" );
		s.addRow( workbook,"3a" );
		expect( s.getColumnCount( workbook ) ).toBe( 3 );
	});

	it( "Can get the maximum number of columns of a sheet specified by number",function() {
		var workbook=s.new();
		s.createSheet( workbook );//add a second sheet and switch to it
		s.setActiveSheetNumber( workbook,2 );
		s.addRow( workbook,"1a,1b" );
		s.addRow( workbook,"2a,2b,2c" );
		s.addRow( workbook,"3a" );
		s.setActiveSheetNumber( workbook,1 );//switch back to sheet 1
		expect( s.getColumnCount( workbook ) ).toBe( 0 );
		expect( s.getColumnCount( workbook,2 ) ).toBe( 3 );
	});

	it( "Can get the maximum number of columns of a sheet specified by name",function() {
		var workbook=s.new();
		s.createSheet( workbook,"test" );
		s.setActiveSheetNumber( workbook,2 );
		s.addRow( workbook,"1a,1b" );
		s.addRow( workbook,"2a,2b,2c" );
		s.addRow( workbook,"3a" );
		s.setActiveSheetNumber( workbook,1 );
		expect( s.getColumnCount( workbook ) ).toBe( 0 );
		expect( s.getColumnCount( workbook,"test" ) ).toBe( 3 );
	});

	describe( "getColumnCount exceptions",function(){

		it( "Throws an exception if the sheet name or number doesn't exist",function() {
			expect( function(){
				var workbook=s.new();
				var result=s.getColumnCount( workbook,2 );
			}).toThrow( regex="Invalid sheet" );
			expect( function(){
				var workbook=s.new();
				var result=s.getColumnCount( workbook,"test" );
			}).toThrow( regex="Invalid sheet" );
		});

	});	

});	
</cfscript>