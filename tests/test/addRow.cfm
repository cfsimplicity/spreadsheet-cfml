<cfscript>
describe( "addRow tests",function(){

	beforeEach( function(){
		variables.data = "a,b";
		variables.workbook = s.new();
	});

	it( "can append a row with the minimum arguments",function() {
		s.addRow( workbook,data );
		s.addRow( workbook,"c,d" );// should be inserted at row 2
		s.write( workbook,tempXlsPath,true );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
		actual = s.read( src=tempXlsPath,format="query" );
		expect( actual ).toBe( expected );
	});

	it( "can insert a row at a specifed position",function() {
		s.addRow( workbook,data );
		s.addRow( workbook,"e,f",2,2 );
		s.addRow( workbook,"c,d",2,1 );
		s.write( workbook,tempXlsPath,true );
		expected = QueryNew( "column1,column2,column3","VarChar,VarChar,VarChar",[ [ "a","b","" ],[ "c","d","" ],[ "","e","f" ] ] );
		actual = s.read( src=tempXlsPath,format="query",includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	it( "can replace a row if insert is false",function() {
		s.addRow( workbook,data );
		s.addRow( workbook=workbook,data=data,startRow=1,insert=false );
		s.write( workbook,tempXlsPath,true );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","b" ] ] );
		actual = s.read( src=tempXlsPath,format="query" );
		expect( actual ).toBe( expected );
	});

});	
</cfscript>