<cfscript>
describe( "workbookFromCsv", function(){

	it( "Returns a workbook from a csv", function(){
		var csv = 'column1,column2#crlf#"Frumpo McNugget",12345';
		var basicExpectedQuery = QueryNew( "column1,column2", "", [ [ "Frumpo McNugget", "12345" ] ] );
		var xls = s.workbookFromCsv( csv=csv, firstRowIsHeader=true );
		var xlsx = s.workbookFromCsv( csv=csv, firstRowIsHeader=true, xmlFormat=true );
		var workbooks = [ xls, xlsx ];
		workbooks.Each( function( wb ){
			actual = s.sheetToQuery( workbook=wb, headerRow=1 );
			expect( actual ).toBe( basicExpectedQuery );
		});
	});

});	
</cfscript>