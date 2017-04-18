<cfscript>
describe( "deleteColumn",function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Deletes the data in a specified column",function() {
		s.addColumn( workbook,"a,b" );
		s.addColumn( workbook,"c,d" );
		s.deleteColumn( workbook,1 );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "","c" ],[ "","d" ] ] );
		actual = s.sheetToQuery( workbook=workbook,includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	describe( "deleteColumn throws an exception if" ,function(){

		it( "column is zero or less",function() {
			expect( function(){
				s.deleteColumn( workbook=workbook,column=0 );
			}).toThrow( regex="Invalid column" );
		});

	});

});	
</cfscript>