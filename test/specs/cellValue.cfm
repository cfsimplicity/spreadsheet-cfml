<cfscript>
describe( "cellValue tests",function(){

	beforeEach( function(){
		variables.workbook = s.new();
	});

	it( "Gets the value from the specified cell",function() {
		data =  QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
		s.addRows( workbook,data );
		expect( s.getCellValue( workbook,2,2 ) ).toBe( "d" );
	});

	it( "Sets the specified cell to the specified string value",function() {
		value="test";
		s.setCellValue( workbook,value,1,1 );
		expected = value;
		actual = s.getCellValue( workbook,1,1 );
		expect( actual ).toBe( expected );
	});

	it( "Sets the specified cell to the specified numeric value",function() {
		value=1;
		s.setCellValue( workbook,value,1,1 );
		expected = value;
		actual = s.getCellValue( workbook,1,1 );
		expect( actual ).toBe( expected );
		expect( IsNumeric( actual ) ).toBeTrue();
	});

	it( "Sets the specified cell to the specified boolean value",function() {
		value=true;
		s.setCellValue( workbook,value,1,1 );
		expected = value;
		actual = s.getCellValue( workbook,1,1 );
		expect( actual ).toBe( expected );
		expect( IsBoolean( actual ) ).toBeTrue();
	});

	it( "Sets the specified cell to the specified date value",function() {
		value=CreateDate( 2015,04,12 );
		s.setCellValue( workbook,value,1,1 );
		expected = DateFormat( value,"yyyy-mm-dd" );
		actual = s.getCellValue( workbook,1,1 );
		expect( actual ).toBe( expected );
	});

	it( "Sets the specified range of cells to the specified value",function() {
		value="a";
		s.setCellRangeValue( workbook,value,1,2,1,2 );
		expected = querySim(
			"column1,column2
			a|a
			a|a");
		s.write( workbook,tempXlsPath,true );
		actual = s.read( src=tempXlsPath,format="query" );
		expect( actual ).toBe( expected );
	});

});	
</cfscript>