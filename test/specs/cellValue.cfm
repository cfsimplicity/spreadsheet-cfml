<cfscript>
describe( "cellValue", function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Gets the value from the specified cell", function() {
		data =  QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
		s.addRows( workbook,data );
		expect( s.getCellValue( workbook, 2, 2 ) ).toBe( "d" );
	});

	it( "Sets the specified cell to the specified string value", function() {
		value = "test";
		s.setCellValue( workbook, value, 1, 1 );
		actual = s.getCellValue( workbook, 1, 1 );
		expect( actual ).toBe( value );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "string" );
	});

	it( "Sets the specified cell to the specified numeric value", function() {
		value = 1;
		s.setCellValue( workbook, value, 1, 1 );
		actual = s.getCellValue( workbook, 1, 1 );
		expect( actual ).toBe( value );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "numeric" );
	});

	it( "Sets the specified cell to the specified date value", function() {
		value = CreateDate( 2015, 04, 12 );
		s.setCellValue( workbook, value, 1, 1 );
		expected = DateFormat( value, "yyyy-mm-dd" );
		actual = s.getCellValue( workbook, 1, 1 );
		expect( actual ).toBe( expected );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "numeric" );
	});

	it( "Sets the specified cell to the specified boolean value with a data type of string by default", function() {
		var value = true;
		s.setCellValue( workbook, value, 1, 1 );
		actual = s.getCellValue( workbook, 1, 1 );
		expect( actual ).toBe( value );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "string" );
	});

	it( "Sets the specified range of cells to the specified value", function() {
		value="a";
		s.setCellRangeValue( workbook, value, 1, 2, 1, 2 );
		expected = querySim(
			"column1,column2
			a|a
			a|a");
		s.write( workbook,tempXlsPath,true );
		actual = s.read( src=tempXlsPath, format="query" );
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

	it( "handles non-date values correctly that Lucee parses as partial dates far in the future", function(){
		value = "01-23112";
		s.setCellValue( workbook, value, 1, 1 );
		expected = "01-23112";
		actual = s.getCellValue( workbook, 1, 1 );
		expect( actual ).toBe( expected );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "string" );
		value = "23112-01";
		s.setCellValue( workbook, value, 1, 1 );
		expected = "23112-01";
		actual = s.getCellValue( workbook, 1, 1 );
		expect( actual ).toBe( expected );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "string" );
	});

	it( "does not accept '9a' or '9p' or '9 a' as valid dates, correcting ACF", function() {
		values = [ "9a", "9p", "9 a", "9    p", "9A" ];
		values.Each( function( value ){
			s.setCellValue( workbook, value, 1, 1 );
			expected = value;
			actual = s.getCellValue( workbook, 1, 1 );
			expect( actual ).toBe( expected );
			expect( s.getCellType( workbook, 1, 1 ) ).toBe( "string" );
		});
	});

	describe( "allows the auto data type detection to be overridden", function(){

		it( "allows forcing values to be added as strings", function(){
			value = 1.234;
			s.setCellValue( workbook, value, 1, 1, "string" );
			actual = s.getCellValue( workbook, 1, 1 );
			expect( actual ).toBe( "1.234" );
			expect( s.getCellType( workbook, 1, 1 ) ).toBe( "string" );
		});

		it( "allows forcing values to be added as numbers", function(){
			value = "0123";
			s.setCellValue( workbook, value, 1, 1, "numeric" );
			actual = s.getCellValue( workbook, 1, 1 );
			expect( actual ).toBe( 123 );
			expect( s.getCellType( workbook, 1, 1 ) ).toBe( "numeric" );
		});

		it( "allows forcing values to be added as dates", function(){
			value = "01.1990";
			s.setCellValue( workbook, value, 1, 1, "date" );
			actual = s.getCellValue( workbook, 1, 1 );
			expect( DateFormat( actual, "yyyy-mm-dd" ) ).toBe( "1990-01-01" );
			expect( s.getCellType( workbook, 1, 1 ) ).toBe( "numeric" );// dates are numeric in Excel
		});

		it( "allows forcing values to be added as booleans", function(){
			values = [ "true", true, 1, "1", "yes", 10 ];
			for( var value in values ){
				s.setCellValue( workbook, value, 1, 1, "boolean" );
				actual = s.getCellValue( workbook, 1, 1 );
				expect( actual ).toBeTrue();
				expect( s.getCellType( workbook, 1, 1 ) ).toBe( "boolean" );
			}
		});

		it( "allows forcing values to be added as blanks", function(){
			values = [ "", "blah" ];
			for( var value in values ){
				s.setCellValue( workbook, value, 1, 1, "blank" );
				actual = s.getCellValue( workbook, 1, 1 );
				expect( actual ).toBeEmpty();
				expect( s.getCellType( workbook, 1, 1 ) ).toBe( "blank" );	
			}
		});

	});

	describe( "setCellValue throws an exception if", function(){

		it( "the data type is invalid", function() {
			expect( function(){
				s.setCellValue( workbook, "test", 1, 1, "blah" );
			}).toThrow( regex="Invalid data type" );
		});

	});

});	
</cfscript>