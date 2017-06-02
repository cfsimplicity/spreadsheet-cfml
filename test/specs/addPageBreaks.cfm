<cfscript>
describe( "addPageBreaks",function(){

	beforeEach( function(){
		var columnData = [ "a", "b", "c" ];
		var rowData = [ columnData, columnData, columnData ];
		variables.data = QueryNew( "c1,c2,c3", "VarChar,VarChar,VarChar", rowData );
		variables.xls = s.workbookFromQuery( data, false );
		variables.xlsx = s.workbookFromQuery( data=data, addHeaderRow=false, xmlformat=true );
		makePublic( s, "getActiveSheet" );
		variables.xlsSheet = s.getActiveSheet( xls );
		variables.xlsxSheet = s.getActiveSheet( xlsx );
	});

	it( "adds page breaks at the row and column numbers passed in as lists",function() {
		s.addPageBreaks( xls, "2,3", "1,2" );
		s.addPageBreaks( xlsx, "2,3", "1,2" );
		expect( xlsSheet.getRowBreaks() ).toBe( [ 1, 2 ] );
		expect( xlsxSheet.getRowBreaks() ).toBe( [ 1, 2 ] );
		expect( xlsSheet.getColumnBreaks() ).toBe( [ 0, 1 ] );
		expect( xlsxSheet.getColumnBreaks() ).toBe( [ 0, 1 ] );
	});

	it( "Doesn't error when passing valid arguments with extra trailing/leading space",function() {
		s.addPageBreaks( xls, " 2,3 ", "1,2 " );
		s.addPageBreaks( xlsx, " 2,3 ", "1,2 " );
	});

	it( "Doesn't error when passing single numbers instead of lists",function() {
		s.addPageBreaks( xls, 1, 2 );
		s.addPageBreaks( xlsx, 1, 2 );
	});

	it( "Throws a helpful exception if both arguments are missing or present but empty",function() {
		expect( function(){
			s.addPageBreaks( xls );
		}).toThrow( regex="Missing argument" );
		expect( function(){
			s.addPageBreaks( xls, "" );
		}).toThrow( regex="Missing argument" );
	});

});	
</cfscript>