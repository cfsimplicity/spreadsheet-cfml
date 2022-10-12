<cfscript>
describe( "addPageBreaks", function(){

	beforeEach( function(){
		var columnData = [ "a", "b", "c" ];
		var rowData = [ columnData, columnData, columnData ];
		var data = QueryNew( "c1,c2,c3", "VarChar,VarChar,VarChar", rowData );
		var xls = s.workbookFromQuery( data, false );
		var xlsx = s.workbookFromQuery( data=data, addHeaderRow=false, xmlformat=true );
		variables.workbooks = [ xls, xlsx ];
	});

	it( "adds page breaks at the row and column numbers passed in as lists", function(){
		workbooks.Each( function( wb ) {
			s.addPageBreaks( wb, "2,3", "1,2" );
			expect( s.getSheetHelper().getActiveSheet( wb ).getRowBreaks() ).toBe( [ 1, 2 ] );
			expect( s.getSheetHelper().getActiveSheet( wb ).getColumnBreaks() ).toBe( [ 0, 1 ] );
		});
	});

	it( "Doesn't error when passing valid arguments with extra trailing/leading space", function(){
		workbooks.Each( function( wb ) {
			s.addPageBreaks( wb, " 2,3 ", "1,2 " );
		});
	});

	it( "Doesn't error when passing single numbers instead of lists", function(){
		workbooks.Each( function( wb ) {
			s.addPageBreaks( wb, 1, 2 );
		});
	});

	it( "Is chainable", function(){
		workbooks.Each( function( wb ) {
			s.newChainable( wb ).addPageBreaks( 1, 2 );
		});
	});

	it( "Throws a helpful exception if both arguments are missing or present but empty", function(){
		workbooks.Each( function( wb ) {
			expect( function(){
				s.addPageBreaks( wb );
			}).toThrow( type="cfsimplicity.spreadsheet.missingRowOrColumnBreaksArgument" );
			expect( function(){
				s.addPageBreaks( wb, "" );
			}).toThrow( type="cfsimplicity.spreadsheet.missingRowOrColumnBreaksArgument" );
		});
	});

});	
</cfscript>