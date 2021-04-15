<cfscript>
describe( "shiftColumns", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Shifts columns right if offset is positive", function(){
		workbooks.Each( function( wb ){
			s.addColumn( wb, "a,b" );
			s.addColumn( wb, "c,d" );
			s.shiftColumns( wb, 1, 1, 1 );
			var expected = querySim( "column1,column2
				|a
				|b
			");
			var actual = s.sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		});
	});

	it( "Shifts columns left if offset is negative", function(){
		workbooks.Each( function( wb ){
			s.addColumn( wb, "a,b" );
			s.addColumn( wb, "c,d" );
			s.addColumn( wb, "e,f" );
			s.shiftColumns( wb, 2, 2, -1 );
			var expected = querySim( "column1,column2,column3
				c||e
				d||f");
			var actual = s.sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		});
	});

});	
</cfscript>