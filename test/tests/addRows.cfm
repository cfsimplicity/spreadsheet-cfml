<cfscript>
describe( "addRows tests",function(){

	beforeEach( function(){
		variables.data = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
		variables.workbook = s.new();
	});

	it( "Appends multiple rows from a query with the minimum arguments",function() {
		s.addRow( workbook,"x,y" );
		s.addRows( workbook,data );
		s.write( workbook,tempXlsPath,true );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "x","y" ],[ "a","b" ],[ "c","d" ] ] );
		actual = s.read( src=tempXlsPath,format="query" );
		expect( actual ).toBe( expected );
	});

	it( "Inserts multiple rows at a specifed position",function() {
		s.addRow( workbook,"e,f" );
		s.addRows( workbook,data,1,2 );
		s.write( workbook,tempXlsPath,true );
		expected = QueryNew( "column1,column2,column3","VarChar,VarChar,VarChar",[ [ "","a","b" ],[ "","c","d" ],[ "e","f","" ] ] );
		actual = s.read( src=tempXlsPath,format="query",includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	it( "Replaces rows if insert is false",function() {
		s.addRow( workbook,"e,f" );
		s.addRow( workbook,"g,h" );
		s.addRows( workbook=workbook,data=data,row=1,insert=false );
		s.write( workbook,tempXlsPath,true );
		expected = data;
		actual = s.read( src=tempXlsPath,format="query" );
		expect( actual ).toBe( expected );
	});

});	
</cfscript>