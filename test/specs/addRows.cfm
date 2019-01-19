<cfscript>
describe( "addRows",function(){

	beforeEach( function(){
		variables.data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
		variables.dataAsArray = [ [ "a", "b" ], [ "c", "d" ] ];
		variables.workbook = s.new();
	});

	it( "Appends multiple rows from a query with the minimum arguments",function() {
		s.addRow( workbook, "x,y" );
		s.addRows( workbook, data );
		expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "x", "y" ], [ "a", "b" ], [ "c", "d" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Can accept data as an array instead of a query", function(){
		var data = [ [ "x", "y" ], [ "a", "b" ], [ "c", "d" ] ];
		s.addRows( workbook, data );
		expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "x", "y" ], [ "a", "b" ], [ "c", "d" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Does nothing if array data is empty",function() {
		workbook = s.new();
		var emptyData = [];
		s.addRows( workbook, emptyData );
		expected = QueryNew( "" );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Inserts multiple rows at a specifed position",function() {
		s.addRow( workbook, "e,f" );
		s.addRows( workbook, data, 1, 2 );
		expected = QueryNew( "column1,column2,column3", "VarChar,VarChar,VarChar", [ [ "", "a", "b" ], [ "", "c", "d" ], [ "e", "f", "" ] ] );
		actual = s.sheetToQuery( workbook=workbook, includeBlankRows=true );
		expect( actual ).toBe( expected );
		//array data
		workbook = s.new();
		s.addRow( workbook, "e,f" );
		s.addRows( workbook, dataAsArray, 1, 2 );
		actual = s.sheetToQuery( workbook=workbook, includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

	it( "Replaces rows if insert is false",function() {
		s.addRow( workbook, "e,f" );
		s.addRow( workbook, "g,h" );
		s.addRows( workbook=workbook, data=data, row=1, insert=false );
		expected = data;
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
		//array data
		workbook = s.new();
		s.addRow( workbook, "e,f" );
		s.addRow( workbook, "g,h" );
		s.addRows( workbook=workbook, data=dataAsArray, row=1, insert=false );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Adds numeric values correctly",function() {
		var data = QueryNew( "column1,column2,column3", "Integer,BigInt,Double", [ [ 1, 1, 1.1 ] ] );
		s.addRows( workbook, data );
		expected = data;
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
		expect( IsNumeric( s.getCellValue( workbook, 1, 1 ) ) ).tobeTrue();
		expect( IsNumeric( s.getCellValue( workbook, 1, 2 ) ) ).tobeTrue();
		expect( IsNumeric( s.getCellValue( workbook, 1, 3 ) ) ).tobeTrue();
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "numeric" );
		expect( s.getCellType( workbook, 1, 2 ) ).toBe( "numeric" );
		expect( s.getCellType( workbook, 1, 3 ) ).toBe( "numeric" );
		//array data
		workbook = s.new();
		var dataAsArray = [ [ 1, 1, 1.1 ] ];
		s.addRows( workbook, dataAsArray );
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
		var data = QueryNew( "column1", "Bit", [ [ true ] ] );
		s.addRows( workbook, data );
		expected = data;
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
		expect( IsBoolean( s.getCellValue( workbook, 1, 1 ) ) ).toBeTrue();
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "boolean" );
		//array data
		workbook = s.new();
		var dataAsArray = [ [ true ] ];
		s.addRows( workbook, dataAsArray );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
		expect( IsBoolean( s.getCellValue( workbook, 1, 1 ) ) ).toBeTrue();
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "string" );// don't set the cell type as boolean from an array
	});

	it( "Adds date/time values correctly",function() {
		var dateValue = CreateDate( 2015, 04, 12 );
		var timeValue = CreateTime( 1, 0, 0 );
		var dateTimeValue = createDateTime( 2015, 04, 12, 1, 0, 0 );
		var data = QueryNew( "column1,column2,column3", "Date,Time,Timestamp",[ [ dateValue, timeValue, dateTimeValue ] ] );
		s.addRows( workbook, data );
		expected = data;
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
		expect( IsDate( s.getCellValue( workbook, 1, 1 ) ) ).toBeTrue();
		expect( IsDate( s.getCellValue( workbook, 1, 2 ) ) ).toBeTrue();
		expect( IsDate( s.getCellValue( workbook, 1, 3 ) ) ).toBeTrue();
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "numeric" );
		expect( s.getCellType( workbook, 1, 2 ) ).toBe( "numeric" );
		expect( s.getCellType( workbook, 1, 3 ) ).toBe( "numeric" );
		//array data
		workbook = s.new();
		var dateValue = CreateDate( 2015, 04, 12 );
		var timeValue = CreateTime( 1, 0, 0 );
		var dateTimeValue = createDateTime( 2015, 04, 12, 1, 0, 0 );
		var dataAsArray = [ [ dateValue, timeValue, dateTimeValue ] ];
		s.addRows( workbook, dataAsArray );
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
		var data = QueryNew( "column1", "Integer", [ [ 0 ] ] );
		s.addRows( workbook, data );
		expected = data;
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
		//array data
		workbook = s.new();
		var dataAsArray = [ [ 0 ] ];
		s.addRows( workbook, dataAsArray );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Handles blank date and boolean values correctly", function(){
		var data = QueryNew( "column1,column2,column3,column4,column5", "Date,Time,Timestamp,Bit,Integer",[ [ "", "", "", "", "" ] ] );
		s.addRows( workbook, data );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "blank" );
		expect( s.getCellType( workbook, 1, 2 ) ).toBe( "blank" );
		expect( s.getCellType( workbook, 1, 3 ) ).toBe( "blank" );
		expect( s.getCellType( workbook, 1, 4 ) ).toBe( "blank" );
		expect( s.getCellType( workbook, 1, 5 ) ).toBe( "numeric" );
		//doesn't apply to array data which has no column types
	});

	it( "Adds strings with leading zeros as strings not numbers",function(){
		var data = QueryNew( "column1", "VarChar", [ [ "01" ] ] );
		s.addRows( workbook, data );
		expected = data;
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
		//array data
		workbook = s.new();
		var dataAsArray = [ [ "01" ] ];
		s.addRows( workbook, dataAsArray );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Can include the query column names",function(){
		s.addRows( workbook=workbook, data=data, includeQueryColumnNames=true );
		expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "column1", "column2" ], [ "a","b" ], [ "c", "d" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Can include the query column names starting at a specific row",function(){
		s.addRow( workbook, "x,y" );
		s.addRows( workbook=workbook, data=data, row=2, includeQueryColumnNames=true );
		expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "x", "y" ], [ "column1", "column2" ], [ "a", "b" ], [ "c", "d" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Can include the query column names starting at a specific column",function(){
		s.addRows( workbook=workbook, data=data, column=2, includeQueryColumnNames=true );
		expected = QueryNew( "column1,column2,column3", "VarChar,VarChar,VarChar", [ [ "", "column1", "column2" ], [ "", "a", "b" ], [ "", "c", "d" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Can include the query column names starting at a specific row and column",function(){
		s.addRow( workbook, "x,y" );
		s.addRows( workbook=workbook, data=data, row=2, column=2, includeQueryColumnNames=true );
		expected = QueryNew( "column1,column2,column3", "VarChar,VarChar,VarChar", [ [ "x", "y", "" ],[ "", "column1","column2" ], [ "", "a", "b" ], [ "", "c","d" ] ] );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	describe( "addRows throws an exception if",function(){

		/* Skip this test by default: can take a long time */
		xit( "adding more than 65536 rows to a binary spreadsheet",function() {
			expect( function(){
				var rows=[];
				for( i=1; i <= 65537; i++ ){
					rows.append( [ i ] );
				}
				var data=QueryNew( "ID","Integer",rows );
				s.addRows( workbook,data );
			}).toThrow( regex="Too many rows" );
		});

		it( "the data is neither a query nor an array", function() {
			expect( function(){
				s.addRows( workbook, "string,list" );
			}).toThrow( message="Invalid data" );
		});

		it( "the data is an array which does not contain an array for each row", function() {
			expect( function(){
				s.addRows( workbook, [ { col1: "a" }, { col2: "b" } ] );// array of structs
			}).toThrow( message="Invalid data" );
		});

	});

});	
</cfscript>