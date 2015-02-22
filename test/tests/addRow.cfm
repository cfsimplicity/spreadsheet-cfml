<cfscript>
describe( "addRow tests",function(){

	beforeEach( function(){
		variables.data = "a,b";
		variables.workbook = s.new();
	});

	it( "Appends a row with the minimum arguments",function() {
		s.addRow( workbook,data );
		s.addRow( workbook,"c,d" );// should be inserted at row 2
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Appends a row including commas with a custom delimiter",function() {
		s.addRow( workbook=workbook,data="a,b|c,d",delimiter="|" );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a,b","c,d" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Inserts a row at a specifed position",function() {
		s.addRow( workbook,data );
		s.addRow( workbook,"e,f",2,2 );
		s.addRow( workbook,"c,d",2,1 );
		expected = QueryNew( "column1,column2,column3","VarChar,VarChar,VarChar",[ [ "a","b","" ],[ "c","d","" ],[ "","e","f" ] ] );
		actual = s.sheetToQuery( workbook=workbook,includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	it( "Replaces a row if insert is false",function() {
		s.addRow( workbook,data );
		s.addRow( workbook=workbook,data=data,row=1,insert=false );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","b" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Handles embedded commas",function() {
		s.addRow( workbook=workbook,data="'a,b','c,d'" );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a,b","c,d" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Adds numeric, boolean or date values correctly",function() {
		var dateValue = CreateDate( 2015,04,12 );
		s.addRow( workbook,"2,true,#dateValue#" );
		expected = QueryNew( "column1,column2,column3","Numeric,Boolean,Date",[ [ 2,true,dateValue ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
		expect( IsNumeric( s.getCellValue( workbook,1,1 ) ) ).tobeTrue();
		expect( IsBoolean( s.getCellValue( workbook,1,2 ) ) ).tobeTrue();
		expect( IsDate( s.getCellValue( workbook,1,3 ) ) ).tobeTrue();
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