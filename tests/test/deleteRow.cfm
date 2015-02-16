<cfscript>
describe( "deleteRow tests",function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "can delete the data in a specified row",function() {
		s.addRow( workbook,"a,b" );
		s.addRow( workbook,"c,d" );
		s.deleteRow( workbook,1 );
		s.write( workbook,tempXlsPath,true );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "","" ],[ "c","d" ] ] );
		actual = s.read( src=tempXlsPath,format="query",includeBlankRows=true );
		expect( actual ).toBe( expected );
	});


});	
</cfscript>