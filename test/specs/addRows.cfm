<cfscript>
describe( "addRows",function(){

	beforeEach( function(){
		variables.rowData = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
		variables.workbook = s.new();
	});

	it( "Appends multiple rows from a query with the minimum arguments",function() {
		s.addRow( workbook,"x,y" );
		s.addRows( workbook,rowData );
		expected = QueryNew( "column1,column2","VarChar,VarChar",[ [ "x","y" ],[ "a","b" ],[ "c","d" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Inserts multiple rows at a specifed position",function() {
		s.addRow( workbook,"e,f" );
		s.addRows( workbook,rowData,1,2 );
		expected = QueryNew( "column1,column2,column3","VarChar,VarChar,VarChar",[ [ "","a","b" ],[ "","c","d" ],[ "e","f","" ] ] );
		actual = s.sheetToQuery( workbook=workbook,includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	it( "Replaces rows if insert is false",function() {
		s.addRow( workbook,"e,f" );
		s.addRow( workbook,"g,h" );
		s.addRows( workbook=workbook,data=rowData,row=1,insert=false );
		expected = rowData;
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Adds numeric values correctly",function() {
		var rowData = QueryNew( "column1,column2,column3", "Integer,BigInt,Double", [ [ 1, 1, 1.1 ] ] );
		s.addRows( workbook, rowData );
		expected = rowData;
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
		expect( IsNumeric( s.getCellValue( workbook, 1, 1 ) ) ).tobeTrue();
		expect( IsNumeric( s.getCellValue( workbook, 1, 2 ) ) ).tobeTrue();
		expect( IsNumeric( s.getCellValue( workbook, 1, 3 ) ) ).tobeTrue();
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "numeric" );
		expect( s.getCellType( workbook, 1, 2 ) ).toBe( "numeric" );
		expect( s.getCellType( workbook, 1, 3 ) ).toBe( "numeric" );
	});

	it( "Adds boolean values correctly",function() {
		var rowData = QueryNew( "column1", "Bit", [ [ true ] ] );
		s.addRows( workbook, rowData );
		expected = rowData;
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
		expect( IsBoolean( s.getCellValue( workbook, 1, 1 ) ) ).toBeTrue();
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "boolean" );
	});

	it( "Adds date/time values correctly",function() {
		var dateValue = CreateDate( 2015, 04, 12 );
		var timeValue = CreateTime( 1, 0, 0 );
		var dateTimeValue = createDateTime( 2015, 04, 12, 1, 0, 0 );
		var rowData = QueryNew( "column1,column2,column3", "Date,Time,Timestamp",[ [ dateValue, timeValue, dateTimeValue ] ] );
		s.addRows( workbook, rowData );
		expected = rowData;
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
		expect( IsDate( s.getCellValue( workbook, 1, 1 ) ) ).toBeTrue();
		expect( IsDate( s.getCellValue( workbook, 1, 2 ) ) ).toBeTrue();
		expect( IsDate( s.getCellValue( workbook, 1, 3 ) ) ).toBeTrue();
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "numeric" );
		expect( s.getCellType( workbook, 1, 2 ) ).toBe( "numeric" );
		expect( s.getCellType( workbook, 1, 3 ) ).toBe( "numeric" );
	});

	it( "Adds zeros as zeros, not booleans",function(){
		var rowData=QueryNew( "column1","Integer",[ [ 0 ] ] );
		s.addRows( workbook,rowData );
		expected=rowData;
		actual=s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Adds strings with leading zeros as strings not numbers",function(){
		var rowData=QueryNew( "column1","VarChar",[ [ "01" ] ] );
		s.addRows( workbook,rowData );
		expected=rowData;
		actual=s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Can include the query column names",function(){
		s.addRows( workbook=workbook, data=rowData, includeQueryColumnNames=true );
		expected=QueryNew( "column1,column2","VarChar,VarChar",[ [ "column1","column2" ],[ "a","b" ],[ "c","d" ] ] );
		actual=s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Can include the query column names starting at a specific row",function(){
		s.addRow( workbook,"x,y" );
		s.addRows( workbook=workbook, data=rowData, row=2, includeQueryColumnNames=true );
		expected=QueryNew( "column1,column2","VarChar,VarChar",[ [ "x","y" ],[ "column1","column2" ],[ "a","b" ],[ "c","d" ] ] );
		actual=s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	describe( "Throws an exception if",function(){

		/* Skip this test by default: can take a long time */
		xit( "adding more than 65536 rows to a binary spreadsheet",function() {
			expect( function(){
				var rows=[];
				for( i=1; i <= 65537; i++ ){
					rows.append( [ i ] );
				}
				var rowData=QueryNew( "ID","Integer",rows );
				s.addRows( workbook,rowData );
			}).toThrow( regex="Too many rows" );
		});

	});

});	
</cfscript>