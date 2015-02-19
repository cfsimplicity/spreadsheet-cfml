<cfscript>
describe( "addRow tests",function(){

	beforeEach( function(){
		variables.data = "a,b";
		variables.workbook = s.new();
	});

	it( "Appends a row with the minimum arguments",function() {
		s.addRow( workbook,data );
		s.addRow( workbook,"c,d" );// should be inserted at row 2
		s.write( workbook,tempXlsPath,true );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
		actual = s.read( src=tempXlsPath,format="query" );
		expect( actual ).toBe( expected );
	});

	it( "Appends a row including commas with a custom delimiter",function() {
		s.addRow( workbook=workbook,data="a,b|c,d",delimiter="|" );
		s.write( workbook,tempXlsPath,true );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a,b","c,d" ] ] );
		actual = s.read( src=tempXlsPath,format="query" );
		expect( actual ).toBe( expected );
	});

	it( "Inserts a row at a specifed position",function() {
		s.addRow( workbook,data );
		s.addRow( workbook,"e,f",2,2 );
		s.addRow( workbook,"c,d",2,1 );
		s.write( workbook,tempXlsPath,true );
		expected = QueryNew( "column1,column2,column3","VarChar,VarChar,VarChar",[ [ "a","b","" ],[ "c","d","" ],[ "","e","f" ] ] );
		actual = s.read( src=tempXlsPath,format="query",includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	it( "Replaces a row if insert is false",function() {
		s.addRow( workbook,data );
		s.addRow( workbook=workbook,data=data,row=1,insert=false );
		s.write( workbook,tempXlsPath,true );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","b" ] ] );
		actual = s.read( src=tempXlsPath,format="query" );
		expect( actual ).toBe( expected );
	});

	it( "Handles embedded commas",function() {
		s.addRow( workbook=workbook,data="'a,b','c,d'" );
		s.write( workbook,tempXlsPath,true );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a,b","c,d" ] ] );
		actual = s.read( src=tempXlsPath,format="query" );
		expect( actual ).toBe( expected );
	});

	describe( "addRow exceptions",function(){

		it( "Throws an exception if row is zero or less",function() {
			expect( function(){
				s.addRow( workbook=workbook,data=data,row=0 );
			}).toThrow( regex="Invalid row" );
		});

		it( "Throws an exception if column is zero or less",function() {
			expect( function(){
				s.addRow( workbook=workbook,data=data,column=0 );
			}).toThrow( regex="Invalid column" );
		});

		it( "Throws an exception if insert is false and no row specified",function() {
			expect( function(){
				s.addRow( workbook=workbook,data=data,insert=false );
			}).toThrow( regex="Missing row" );
		});

	});

});	
</cfscript>