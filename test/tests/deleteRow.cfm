<cfscript>
describe( "deleteRow tests",function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Deletes the data in a specified row",function() {
		s.addRow( workbook,"a,b" );
		s.addRow( workbook,"c,d" );
		s.deleteRow( workbook,1 );
		s.write( workbook,tempXlsPath,true );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "","" ],[ "c","d" ] ] );
		actual = s.read( src=tempXlsPath,format="query",includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	describe( "deleteRow exceptions",function(){

		it( "Throws an exception if row is zero or less",function() {
			expect( function(){
				s.deleteRow( workbook=workbook,row=0 );
			}).toThrow( regex="Invalid row" );
		});

	});

});	
</cfscript>