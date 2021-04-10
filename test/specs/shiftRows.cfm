<cfscript>
describe( "shiftRows", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		variables.rowData = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
	});

	it( "Shifts rows down if offset is positive", function(){
		workbooks.Each( function( wb ){
			s.addRows( wb, rowData );
			s.shiftRows( wb, 1, 1, 1 );
			var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "" ], [ "a", "b" ] ] );
			var actual = s.sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		});
	});

	it( "Shifts rows up if offset is negative", function(){
		workbooks.Each( function( wb ){
			s.addRows( wb, rowData );
			s.shiftRows( wb, 2, 2, -1 );
			var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "c", "d" ] ] );
			var actual = s.sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

});	
</cfscript>