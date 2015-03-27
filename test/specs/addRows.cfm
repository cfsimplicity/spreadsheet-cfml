<cfscript>
describe( "addRows tests",function(){

	beforeEach( function(){
		variables.data = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
		variables.workbook = s.new();
	});

	it( "Appends multiple rows from a query with the minimum arguments",function() {
		s.addRow( workbook,"x,y" );
		s.addRows( workbook,data );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "x","y" ],[ "a","b" ],[ "c","d" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Inserts multiple rows at a specifed position",function() {
		s.addRow( workbook,"e,f" );
		s.addRows( workbook,data,1,2 );
		expected = QueryNew( "column1,column2,column3","VarChar,VarChar,VarChar",[ [ "","a","b" ],[ "","c","d" ],[ "e","f","" ] ] );
		actual = s.sheetToQuery( workbook=workbook,includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	it( "Replaces rows if insert is false",function() {
		s.addRow( workbook,"e,f" );
		s.addRow( workbook,"g,h" );
		s.addRows( workbook=workbook,data=data,row=1,insert=false );
		expected = data;
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Adds numeric, boolean or date values correctly",function() {
		var dateValue = CreateDate( 2015,04,12 );
		var data = QueryNew( "column1,column2,column3","Numeric,Boolean,Date",[ [ 2,true,dateValue ] ] );
		s.addRows( workbook,data );
		expected = data;
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
		expect( IsNumeric( s.getCellValue( workbook,1,1 ) ) ).tobeTrue();
		expect( IsBoolean( s.getCellValue( workbook,1,2 ) ) ).tobeTrue();
		expect( IsDate( s.getCellValue( workbook,1,3 ) ) ).tobeTrue();
	});

});	
</cfscript>