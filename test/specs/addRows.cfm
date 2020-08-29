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
		var dateTimeValue = CreateDateTime( 2015, 04, 12, 1, 0, 0 );
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

	it( "Formats time and timestamp values correctly when custom mask includes fractions of a second",function() {
		dateFormats = {
			TIME: "hh:mm:ss.000"
			,TIMESTAMP: "yyyy-mm-dd hh:mm:ss.000"
		};
		var s = newSpreadsheetInstance( dateFormats: dateFormats );
		/*
		ACF doesn't support milliseconds, ie:
			var timeValue = CreateTime( 1, 0, 0, 999 );
			var dateTimeValue = CreateDateTime( 2015, 04, 12, 1, 0, 0, 999 );
		So use java to create datetime objects including milliseconds
		*/
		var timeValue = CreateObject( "java", "java.util.Date" ).init( JavaCast( "long", 360000999 ) );
		var dateTimeValue = CreateObject( "java", "java.util.Date" ).init( JavaCast( "long", 1428796800999 ) );
		var data = QueryNew( "column1,column2", "Time,Timestamp", [ [ timeValue, dateTimeValue ] ] );
		s.addRows( variables.workbook, data );
		expectedTimeValue = data.column1[ 1 ].TimeFormat( "hh:nn:ss:l" );
		expectedDateTimeValue = data.column2[ 1 ].DateTimeFormat( "yyyy-mm-dd hh:nn:ss:l" );
		actual = s.sheetToQuery( workbook );
		actualTimeValue = actual.column1[ 1 ];
		actualDateTimeValue = actual.column2[ 1 ];
		//array data
		var workbook = s.new();
		var dataAsArray = [ [ timeValue, dateTimeValue ] ];
		s.addRows( workbook, dataAsArray );
		expectedTimeValue = data.column1[ 1 ].TimeFormat( "hh:nn:ss:l" );
		expectedDateTimeValue = data.column2[ 1 ].DateTimeFormat( "yyyy-mm-dd hh:nn:ss:l" );
		actual = s.sheetToQuery( workbook );
		actualTimeValue = actual.column1[ 1 ];
		actualDateTimeValue = actual.column2[ 1 ];
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

	it( "Handles empty values correctly", function(){
		var data = QueryNew( "column1,column2,column3,column4,column5", "Date,Time,Timestamp,Bit,Integer",[ [ "", "", "", "", "" ] ] );
		s.addRows( workbook, data );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "blank" );
		expect( s.getCellType( workbook, 1, 2 ) ).toBe( "blank" );
		expect( s.getCellType( workbook, 1, 3 ) ).toBe( "blank" );
		expect( s.getCellType( workbook, 1, 4 ) ).toBe( "blank" );
		expect( s.getCellType( workbook, 1, 5 ) ).toBe( "numeric" );
		//doesn't apply to array data which has no column types
	});

	it( "Can ignore query column types, so that each cell's type is auto-detected from its value", function(){
		var dateValue = CreateDate( 2015, 04, 12 );
		var data = QueryNew( "column1", "VarChar", [ [ 0 ], [ 1 ], [ 1.1 ], [ dateValue ], [ "hello" ] ] );
		s.addRows( workbook=workbook, data=data, ignoreQueryColumnDataTypes=true );
		expect( s.getCellType( workbook, 1, 1 ) ).toBe( "numeric" );
		expect( s.getCellType( workbook, 2, 1 ) ).toBe( "numeric" );
		expect( s.getCellType( workbook, 3, 1 ) ).toBe( "numeric" );
		expect( s.getCellType( workbook, 4, 1 ) ).toBe( "numeric" );
		expect( s.getCellType( workbook, 5, 1 ) ).toBe( "string" );
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
		expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "column1", "column2" ], [ "a","b" ], [ "c", "d" ] ] );
		s.addRows( workbook=workbook, data=data, includeQueryColumnNames=true );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
		//test xlsx
		var workbook = s.newXlsx();
		s.addRows( workbook=workbook, data=data, includeQueryColumnNames=true );
		actual = s.sheetToQuery( workbook );
		expect( actual ).toBe( expected );
	});

	it( "Includes query columns in the same case and order as the original query", function() {
		local.data = QueryNew( "Header2,Header1", "VarChar,VarChar", [ [ "b","a" ], [ "d","c" ] ] );
		s.addRows( workbook=workbook, data=local.data, includeQueryColumnNames=true );
		expect( s.getCellValue( workbook, 1, 1 ) ).toBeWithCase( "Header2" );
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

	it( "Doesn't error if the workbook is SXSSF and autoSizeColumns is true", function(){
		var workbook = s.newStreamingXlsx();
		s.addRows( workbook=local.workbook, data=data, autoSizeColumns=true );
	});

	describe( "addRows() data type overriding",function(){

		it( "throws an error if invalid types are specified in the datatype struct", function() {
			expect( function(){
				var data = [ [ "a", "b" ] ];
				var datatypes = { numeric: [ 1 ], varchar: [ 2 ] };
				s.addRows( workbook=workbook, data=data, datatypes=datatypes );
			}).toThrow( message="Invalid datatype(s)" );
		});

		it( "throws an error if columns to override are not specified as arrays in the datatype struct", function() {
			expect( function(){
				var data = [ [ "a", "b" ] ];
				var datatypes = { numeric: "1", string: "2" };
				s.addRows( workbook=workbook, data=data, datatypes=datatypes );
			}).toThrow( message="Invalid datatype(s)" );
		});

		it( "Allows column data types in data passed as an array to be overridden by column number", function(){
			var data = [ [ "01234", 1234567890123456 ] ];
			s.addRows( workbook, data );
			expect( s.getCellValue( workbook, 1, 1 ) ).toBe( "01234" );
			expect( s.getCellType( workbook, 1, 1 ) ).toBe( "string" );
			expect( s.getCellType( workbook, 1, 2 ) ).toBe( "numeric" );
			var datatypes = { numeric: [ 1 ], string: [ 2 ] };// can't test dates: date strings are always converted correctly!
			s.addRows( workbook=workbook, data=data, datatypes=datatypes );
			expect( s.getCellValue( workbook, 2, 1 ) ).toBe( "1234" );
			expect( s.getCellType( workbook, 2, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( workbook, 2, 2 ) ).toBe( "string" );
		});

		it( "Allows column data types in data passed as a query to be overridden by column name or number", function(){
			var data = QueryNew( "Number,Date,String", "VarChar,VarChar,BigInt", [ [ "01234", "2020-08-24", 1234567890123456 ] ] );
			s.addRows( workbook, data );
			expect( s.getCellValue( workbook, 1, 1 ) ).toBe( "01234" );
			expect( s.getCellType( workbook, 1, 1 ) ).toBe( "string" );
			expect( s.getCellType( workbook, 1, 2 ) ).toBe( "string" );
			expect( s.getCellType( workbook, 1, 3 ) ).toBe( "numeric" );
			var datatypes = { numeric: [ "Number" ], date: [ "Date" ], string: [ "String" ] };
			s.addRows( workbook=workbook, data=data, datatypes=datatypes );
			expect( s.getCellValue( workbook, 2, 1 ) ).toBe( "1234" );
			expect( s.getCellType( workbook, 2, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( workbook, 2, 2 ) ).toBe( "numeric" );//dates are stored as numbers in Excel
			expect( IsDate( s.getCellValue( workbook, 2, 2 ) ) ).toBeTrue();
			expect( s.getCellType( workbook, 2, 3 ) ).toBe( "string" );
			// mixture of column names and numbers
			var data = QueryNew( "Number1,Number2", "VarChar,VarChar", [ [ "01234", "01234" ] ] );
			var datatypes = { numeric: [ "Number1", 2 ] };
			s.addRows( workbook=workbook, data=data, datatypes=datatypes );
			expect( s.getCellValue( workbook, 3, 1 ) ).toBe( "1234" );
			expect( s.getCellValue( workbook, 3, 2 ) ).toBe( "1234" );
			expect( s.getCellType( workbook, 3, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( workbook, 3, 2 ) ).toBe( "numeric" );
		});

		it( "Values in array data fall back to the autodetected type if they don't match the overridden type", function() {
			var data = [ [ "01234", "alpha" ] ];
			var datatypes = { numeric: [ 1, 2 ] };
			s.addRows( workbook=workbook, data=data, datatypes=datatypes );
			expect( s.getCellValue( workbook, 1, 1 ) ).toBe( 1234 );
			expect( s.getCellType( workbook, 1, 1 ) ).toBe( "numeric" );
			expect( s.getCellValue( workbook, 1, 2 ) ).toBe( "alpha" );
			expect( s.getCellType( workbook, 1, 2 ) ).toBe( "string" );
		});

		it( "Values in query data fall back to the query column type if they don't match the overridden type", function() {
			var data = QueryNew( "Number,String,Date", "VarChar,VarChar,VarChar", [ [ "01234", "alpha", "alpha" ] ] );
			var datatypes = { numeric: [ 1, 2 ], date: [ 3 ] };
			s.addRows( workbook=workbook, data=data, datatypes=datatypes );
			expect( s.getCellValue( workbook, 1, 1 ) ).toBe( 1234 );
			expect( s.getCellType( workbook, 1, 1 ) ).toBe( "numeric" );
			expect( s.getCellValue( workbook, 1, 2 ) ).toBe( "alpha" );
			expect( s.getCellType( workbook, 1, 2 ) ).toBe( "string" );
			expect( s.getCellValue( workbook, 1, 3 ) ).toBe( "alpha" );
			expect( s.getCellType( workbook, 1, 3 ) ).toBe( "string" );
		});

		it( "Query data values with NO type override, default to query column types", function() {
			var data = QueryNew( "Number,String", "VarChar,VarChar", [ [ 1234, "01234" ] ] );
			var datatypes = { numeric: [ 2 ] };
			s.addRows( workbook=workbook, data=data, datatypes=datatypes );
			expect( s.getCellType( workbook, 1, 1 ) ).toBe( "string" );
			expect( s.getCellType( workbook, 1, 2 ) ).toBe( "numeric" );
		});

		it( "Values in query data fall back to the autodetected type if they don't match the overridden type and ignoreQueryColumnDataTypes is true", function() {
			var data = QueryNew( "Number,String,Date", "VarChar,VarChar,VarChar", [ [ "01234", "alpha", "alpha" ] ] );
			var datatypes = { numeric: [ 1, 2 ], date: [ 3 ] };
			s.addRows( workbook=workbook, data=data, ignoreQueryColumnDataTypes=true, datatypes=datatypes );
			expect( s.getCellValue( workbook, 1, 1 ) ).toBe( 1234 );
			expect( s.getCellType( workbook, 1, 1 ) ).toBe( "numeric" );
			expect( s.getCellValue( workbook, 1, 2 ) ).toBe( "alpha" );
			expect( s.getCellType( workbook, 1, 2 ) ).toBe( "string" );
			expect( s.getCellValue( workbook, 1, 3 ) ).toBe( "alpha" );
			expect( s.getCellType( workbook, 1, 3 ) ).toBe( "string" );
		});

		it( "Query data values in columns with an override type of 'auto' will have their type auto-detected, regardless of the query column type", function() {
			var data = QueryNew( "One,Two", "VarChar,VarChar", [ [ "2020-08-24", "2020-08-24" ], [ "3.1", "3.1" ] ] );
			var datatypes = { auto: [ 1 ] };
			s.addRows( workbook=workbook, data=data, datatypes=datatypes );
			expect( s.getCellType( workbook, 1, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( workbook, 1, 2 ) ).toBe( "string" );
			expect( s.getCellType( workbook, 2, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( workbook, 2, 2 ) ).toBe( "string" );
		});

	});

	describe( "addRows throws an exception if",function(){

		/* Skip this test by default: can take a long time */
		xit( "adding more than 65536 rows to a binary spreadsheet",function() {
			expect( function(){
				var rows=[];
				for( var i=1; i <= 65537; i++ ){
					rows.append( [ i ] );
				}
				var data=QueryNew( "ID","Integer",rows );
				variables.s.addRows( workbook,data );
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