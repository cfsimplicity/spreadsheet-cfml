<cfscript>
describe( "cellValue",function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Gets the value from the specified cell",function() {
		data =  QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
		s.addRows( workbook,data );
		expect( s.getCellValue( workbook, 2, 2 ) ).toBe( "d" );
	});

	it( "Sets the specified cell to the specified string value",function() {
		value = "test";
		s.setCellValue( workbook, value, 1, 1 );
		actual = s.getCellValue( workbook, 1, 1 );
		expect( actual ).toBe( value );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "string" );
	});

	it( "Sets the specified cell to the specified numeric value",function() {
		value = 1;
		s.setCellValue( workbook, value, 1, 1 );
		actual = s.getCellValue( workbook, 1, 1 );
		expect( actual ).toBe( value );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "numeric" );
	});

	it( "Sets the specified cell to the specified date value",function() {
		value = CreateDate( 2015, 04, 12 );
		s.setCellValue( workbook, value, 1, 1 );
		expected = DateFormat( value, "yyyy-mm-dd" );
		actual = s.getCellValue( workbook, 1, 1 );
		expect( actual ).toBe( expected );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "numeric" );
	});

	it( "Sets the specified cell to the specified boolean value with a data type of string by default",function() {
		var value = true;
		s.setCellValue( workbook, value, 1, 1 );
		actual = s.getCellValue( workbook, 1, 1 );
		expect( actual ).toBe( value );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "string" );
	});

	it( "Sets the specified range of cells to the specified value",function() {
		value="a";
		s.setCellRangeValue( workbook, value, 1, 2, 1, 2 );
		expected = querySim(
			"column1,column2
			a|a
			a|a");
		s.write( workbook,tempXlsPath,true );
		actual = s.read( src=tempXlsPath,format="query" );
		expect( actual ).toBe( expected );
	});

	it( "handles numbers with leading zeros correctly", function(){
		value = "0162220494";
		s.setCellValue( workbook, value, 1, 1 );
		expected = "0162220494";
		actual = s.getCellValue( workbook, 1, 1 );
		expect( actual ).toBe( expected );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "string" );
	});

	it( "handles non-date values correctly that Lucee incorrectly treats as dates", function(){
		value = "01-23112";
		s.setCellValue( workbook, value, 1, 1 );
		expected = "01-23112";
		actual = s.getCellValue( workbook, 1, 1 );
		expect( actual ).toBe( expected );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "string" );
	});

	/* describe( "setCellValue throws an exception if",function(){

		it( "the data type is invalid",function() {
			expect( function(){
				s.setCellValue( workbook, "test", 1, 1, "blah" );
			}).toThrow( regex="Invalid data type" );
		});

	}); */

});	
</cfscript>