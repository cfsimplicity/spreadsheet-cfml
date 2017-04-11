<cfscript>
describe( "workbookFromCsv",function(){

	beforeEach( function(){
		savecontent variable="variables.csv"{
			WriteOutput( '
column1,column2
"Frumpo McNugget",12345
		');
		};
		variables.basicExpectedQuery = QueryNew( "column1,column2", "", [ [ "Frumpo McNugget", "12345" ] ] );
	});

	it( "Returns a workbook from a csv",function() {
		workbook = s.workbookFromCsv( csv=csv, firstRowIsHeader=true );
		actual = s.sheetToQuery( workbook=workbook, headerRow=1 );
		expect( actual ).toBe( basicExpectedQuery );
	});

});	
</cfscript>