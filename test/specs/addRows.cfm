<cfscript>
describe( "addRows", function(){

	beforeEach( function(){
		variables.data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
		variables.dataAsArray = [ [ "a", "b" ], [ "c", "d" ] ];
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Appends multiple rows from a query with the minimum arguments", function(){
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "x", "y" ], [ "a", "b" ], [ "c", "d" ] ] );
		workbooks.Each( function( wb ){
			s.addRow( wb, "x,y" )
				.addRows( wb, data );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Is chainable", function(){
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "x", "y" ], [ "a", "b" ], [ "c", "d" ] ] );
		workbooks.Each( function( wb ){
			s.newChainable( wb )
				.addRow( "x,y" )
				.addRows( data );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Can accept data as an array instead of a query", function(){
		var data = [ [ "x", "y" ], [ "a", "b" ], [ "c", "d" ] ];
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "x", "y" ], [ "a", "b" ], [ "c", "d" ] ] );
		workbooks.Each( function( wb ){
			s.addRows( wb, data );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Does nothing if array data is empty", function(){
		var emptyData = [];
		workbooks.Each( function( wb ){
			s.addRows( wb, emptyData );
			var expected = QueryNew( "" );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Inserts multiple rows at a specifed position", function(){
		var expected = QueryNew( "column1,column2,column3", "VarChar,VarChar,VarChar", [ [ "", "a", "b" ], [ "", "c", "d" ], [ "e", "f", "" ] ] );
		workbooks.Each( function( wb ){
			s.addRow( wb, "e,f" )
				.addRows( wb, data, 1, 2 );
			var actual = s.getSheetHelper().sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		});
		//array data
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			s.addRow( wb, "e,f" )
				.addRows( wb, dataAsArray, 1, 2 );
			actual = s.getSheetHelper().sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		});
	});

	it( "Replaces rows if insert is false", function(){
		var expected = data;
		workbooks.Each( function( wb ){
			s.addRow( wb, "e,f" )
				.addRow( wb, "g,h" )
				.addRows( workbook=wb, data=data, row=1, insert=false );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			s.addRow( wb, "e,f" )
				.addRow( wb, "g,h" )
				.addRows( workbook=wb, data=dataAsArray, row=1, insert=false );
			actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Adds numeric values correctly", function(){
		var data = QueryNew( "column1,column2,column3", "Integer,BigInt,Double", [ [ 1, 1, 1.1 ] ] );
		var expected = data;
		workbooks.Each( function( wb ){
			s.addRows( wb, data );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
			expect( IsNumeric( s.getCellValue( wb, 1, 1 ) ) ).tobeTrue();
			expect( IsNumeric( s.getCellValue( wb, 1, 2 ) ) ).tobeTrue();
			expect( IsNumeric( s.getCellValue( wb, 1, 3 ) ) ).tobeTrue();
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 1, 2 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 1, 3 ) ).toBe( "numeric" );
		});
		var dataAsArray = [ [ 1, 1, 1.1 ] ];
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			s.addRows( wb, dataAsArray );
			actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
			expect( IsNumeric( s.getCellValue( wb, 1, 1 ) ) ).tobeTrue();
			expect( IsNumeric( s.getCellValue( wb, 1, 2 ) ) ).tobeTrue();
			expect( IsNumeric( s.getCellValue( wb, 1, 3 ) ) ).tobeTrue();
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 1, 2 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 1, 3 ) ).toBe( "numeric" );
		});
	});

	it( "Adds boolean values correctly", function(){
		var data = QueryNew( "column1", "Bit", [ [ true ] ] );
		var expected = data;
			workbooks.Each( function( wb ){
			s.addRows( wb, data );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
			expect( IsBoolean( s.getCellValue( wb, 1, 1 ) ) ).toBeTrue();
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "boolean" );
		});
		var dataAsArray = [ [ true ] ];
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			s.addRows( wb, dataAsArray );
			actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
			expect( IsBoolean( s.getCellValue( wb, 1, 1 ) ) ).toBeTrue();
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );// don't set the cell type as boolean from an array
		});
	});

	it( "Adds date/time values correctly", function(){
		var dateValue = CreateDate( 2015, 04, 12 );
		var timeValue = CreateTime( 1, 0, 0 );
		var dateTimeValue = createDateTime( 2015, 04, 12, 1, 0, 0 );
		var data = QueryNew( "column1,column2,column3", "Date,Time,Timestamp", [ [ dateValue, timeValue, dateTimeValue ] ] );
		var expected = data;
		workbooks.Each( function( wb ){
			s.addRows( wb, data );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
			expect( IsDate( s.getCellValue( wb, 1, 1 ) ) ).toBeTrue();
			expect( IsDate( s.getCellValue( wb, 1, 2 ) ) ).toBeTrue();
			expect( IsDate( s.getCellValue( wb, 1, 3 ) ) ).toBeTrue();
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 1, 2 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 1, 3 ) ).toBe( "numeric" );
		});
		//array data
		var dateValue = CreateDate( 2015, 04, 12 );
		var timeValue = CreateTime( 1, 0, 0 );
		var dateTimeValue = CreateDateTime( 2015, 04, 12, 1, 0, 0 );
		var dataAsArray = [ [ dateValue, timeValue, dateTimeValue ] ];
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			s.addRows( wb, dataAsArray );
			actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
			expect( IsDate( s.getCellValue( wb, 1, 1 ) ) ).toBeTrue();
			expect( IsDate( s.getCellValue( wb, 1, 2 ) ) ).toBeTrue();
			expect( IsDate( s.getCellValue( wb, 1, 3 ) ) ).toBeTrue();
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 1, 2 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 1, 3 ) ).toBe( "numeric" );
		});
	});

	it( "Formats time and timestamp values correctly when custom mask includes fractions of a second", function(){
		var dateFormats = {
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
		var expectedTimeValue = data.column1[ 1 ].TimeFormat( "hh:nn:ss:l" );
		var expectedDateTimeValue = data.column2[ 1 ].DateTimeFormat( "yyyy-mm-dd hh:nn:ss:l" );
		workbooks.Each( function( wb ){
			s.addRows( wb, data );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			var actualTimeValue = actual.column1[ 1 ];
			var actualDateTimeValue = actual.column2[ 1 ];
		});
		var dataAsArray = [ [ timeValue, dateTimeValue ] ];
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			s.addRows( wb, dataAsArray );
			expectedTimeValue = data.column1[ 1 ].TimeFormat( "hh:nn:ss:l" );
			expectedDateTimeValue = data.column2[ 1 ].DateTimeFormat( "yyyy-mm-dd hh:nn:ss:l" );
			actual = s.getSheetHelper().sheetToQuery( wb );
			actualTimeValue = actual.column1[ 1 ];
			actualDateTimeValue = actual.column2[ 1 ];
		});
	});

	it( "Adds zeros as zeros, not booleans", function(){
		var data = QueryNew( "column1", "Integer", [ [ 0 ] ] );
		var expected = data;
		workbooks.Each( function( wb ){
			s.addRows( wb, data );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
		var dataAsArray = [ [ 0 ] ];
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			//array data
			s.addRows( wb, dataAsArray );
			actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Handles empty values correctly", function(){
		var data = QueryNew( "column1,column2,column3,column4,column5", "Date,Time,Timestamp,Bit,Integer",[ [ "", "", "", "", "" ] ] );
		workbooks.Each( function( wb ){
			s.addRows( wb, data );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "blank" );
			expect( s.getCellType( wb, 1, 2 ) ).toBe( "blank" );
			expect( s.getCellType( wb, 1, 3 ) ).toBe( "blank" );
			expect( s.getCellType( wb, 1, 4 ) ).toBe( "blank" );
			expect( s.getCellType( wb, 1, 5 ) ).toBe( "numeric" );
			//doesn't apply to array data which has no column types
		});
	});

	it( "Can ignore query column types, so that each cell's type is auto-detected from its value", function(){
		var dateValue = CreateDate( 2015, 04, 12 );
		var data = QueryNew( "column1", "VarChar", [ [ 0 ], [ 1 ], [ 1.1 ], [ dateValue ], [ "hello" ] ] );
		workbooks.Each( function( wb ){
			s.addRows( workbook=wb, data=data, ignoreQueryColumnDataTypes=true );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 2, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 3, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 4, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 5, 1 ) ).toBe( "string" );
		});
	});

	it( "Adds strings with leading zeros as strings not numbers", function(){
		var data = QueryNew( "column1", "VarChar", [ [ "01" ] ] );
		var expected = data;
		workbooks.Each( function( wb ){
			s.addRows( wb, data );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
		var dataAsArray = [ [ "01" ] ];
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			s.addRows( wb, dataAsArray );
			actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Can include the query column names", function(){
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "column1", "column2" ], [ "a","b" ], [ "c", "d" ] ] );
		workbooks.Each( function( wb ){
			s.addRows( workbook=wb, data=data, includeQueryColumnNames=true );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Includes query columns in the same case and order as the original query", function(){
		var data = QueryNew( "Header2,Header1", "VarChar,VarChar", [ [ "b","a" ], [ "d","c" ] ] );
		workbooks.Each( function( wb ){
			s.addRows( workbook=wb, data=data, includeQueryColumnNames=true );
			expect( s.getCellValue( wb, 1, 1 ) ).toBeWithCase( "Header2" );
		});
	});

	it( "Can include the query column names starting at a specific row", function(){
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "x", "y" ], [ "column1", "column2" ], [ "a", "b" ], [ "c", "d" ] ] );
		workbooks.Each( function( wb ){
			s.addRow( wb, "x,y" )
				.addRows( workbook=wb, data=data, row=2, includeQueryColumnNames=true );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Can include the query column names starting at a specific column", function(){
		var expected = QueryNew( "column1,column2,column3", "VarChar,VarChar,VarChar", [ [ "", "column1", "column2" ], [ "", "a", "b" ], [ "", "c", "d" ] ] );
		workbooks.Each( function( wb ){
			s.addRows( workbook=wb, data=data, column=2, includeQueryColumnNames=true );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Can include the query column names starting at a specific row and column", function(){
		var expected = QueryNew( "column1,column2,column3", "VarChar,VarChar,VarChar", [ [ "x", "y", "" ],[ "", "column1","column2" ], [ "", "a", "b" ], [ "", "c","d" ] ] );
		workbooks.Each( function( wb ){
			s.addRow( wb, "x,y" )
				.addRows( workbook=wb, data=data, row=2, column=2, includeQueryColumnNames=true );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Doesn't error if the workbook is SXSSF and autoSizeColumns is true", function(){
		var wb = s.newStreamingXlsx();
		s.addRows( workbook=local.wb, data=data, autoSizeColumns=true );
	});

	describe( "addRows() data type overriding", function(){

		it( "throws an error if invalid types are specified in the datatype struct", function(){
			var data = [ [ "a", "b" ] ];
			var datatypes = { numeric: [ 1 ], varchar: [ 2 ] };
			workbooks.Each( function( wb ){
				expect( function(){
					s.addRows( workbook=wb, data=data, datatypes=datatypes );
				}).toThrow( regex="Invalid datatype\(s\)" );
			});
		});

		it( "throws an error if columns to override are not specified as arrays in the datatype struct", function(){
			var data = [ [ "a", "b" ] ];
			var datatypes = { numeric: "1", string: "2" };
			workbooks.Each( function( wb ){
				expect( function(){
					s.addRows( workbook=wb, data=data, datatypes=datatypes );
				}).toThrow( regex="Invalid datatype\(s\)" );
			});
		});

		it( "Allows column data types in data passed as an array to be overridden by column number", function(){
			var data = [ [ "01234", 1234567890123456 ] ];
			var datatypes = { numeric: [ 1 ], string: [ 2 ] };// can't test dates: date strings are always converted correctly!
			workbooks.Each( function( wb ){
				s.addRows( wb, data );
				expect( s.getCellValue( wb, 1, 1 ) ).toBe( "01234" );
				expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
				expect( s.getCellType( wb, 1, 2 ) ).toBe( "numeric" );
				s.addRows( workbook=wb, data=data, datatypes=datatypes );
				expect( s.getCellValue( wb, 2, 1 ) ).toBe( "1234" );
				expect( s.getCellType( wb, 2, 1 ) ).toBe( "numeric" );
				expect( s.getCellType( wb, 2, 2 ) ).toBe( "string" );
			});
		});

		it( "Allows column data types in data passed as a query to be overridden by column name or number", function(){
			workbooks.Each( function( wb ){
				var data = QueryNew( "Number,Date,String,Time,Boolean", "VarChar,VarChar,BigInt,VarChar,VarChar", [ [ "01234", "2020-08-24", 1234567890123456, "2020-08-24 09:15:00", "yes" ] ] );
				s.addRows( wb, data );
				expect( s.getCellValue( wb, 1, 1 ) ).toBe( "01234" );
				expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
				expect( s.getCellType( wb, 1, 2 ) ).toBe( "string" );
				expect( s.getCellType( wb, 1, 3 ) ).toBe( "numeric" );
				expect( s.getCellType( wb, 1, 4 ) ).toBe( "string" );
				expect( s.getCellType( wb, 1, 5 ) ).toBe( "string" );
				var datatypes = { numeric: [ "Number" ], date: [ "Date" ], string: [ "String" ], time: [ "Time" ], boolean: [ "boolean" ] };
				s.addRows( workbook=wb, data=data, datatypes=datatypes );
				expect( s.getCellValue( wb, 2, 1 ) ).toBe( "1234" );
				expect( s.getCellType( wb, 2, 1 ) ).toBe( "numeric" );
				expect( s.getCellType( wb, 2, 2 ) ).toBe( "numeric" );//dates are stored as numbers in Excel
				expect( IsDate( s.getCellValue( wb, 2, 2 ) ) ).toBeTrue();
				expect( s.getCellType( wb, 2, 3 ) ).toBe( "string" );
				expect( s.getCellType( wb, 2, 4 ) ).toBe( "numeric" );
				expect( s.getCellValue( wb, 2, 4 ) ).toBe( "09:15:00" );
				expect( s.getCellType( wb, 2, 5 ) ).toBe( "boolean" );
				// mixture of column names and numbers
				var data = QueryNew( "Number1,Number2", "VarChar,VarChar", [ [ "01234", "01234" ] ] );
				var datatypes = { numeric: [ "Number1", 2 ] };
				s.addRows( workbook=wb, data=data, datatypes=datatypes );
				expect( s.getCellValue( wb, 3, 1 ) ).toBe( "1234" );
				expect( s.getCellValue( wb, 3, 2 ) ).toBe( "1234" );
				expect( s.getCellType( wb, 3, 1 ) ).toBe( "numeric" );
				expect( s.getCellType( wb, 3, 2 ) ).toBe( "numeric" );
			});
		});

		it( "Values in array data fall back to the autodetected type if they don't match the overridden type", function(){
			var data = [ [ "01234", "alpha", "alpha", "alpha", "alpha" ] ];
			var datatypes = { numeric: [ 1, 2 ], date: [ 3 ], time: [ 4 ], boolean: [ 5 ] };
			workbooks.Each( function( wb ){
				s.addRows( workbook=wb, data=data, datatypes=datatypes );
				expect( s.getCellValue( wb, 1, 1 ) ).toBe( 1234 );
				expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
				expect( s.getCellValue( wb, 1, 2 ) ).toBe( "alpha" );
				expect( s.getCellType( wb, 1, 2 ) ).toBe( "string" );
				expect( s.getCellValue( wb, 1, 3 ) ).toBe( "alpha" );
				expect( s.getCellType( wb, 1, 3 ) ).toBe( "string" );
				expect( s.getCellValue( wb, 1, 4 ) ).toBe( "alpha" );
				expect( s.getCellType( wb, 1, 4 ) ).toBe( "string" );
				expect( s.getCellValue( wb, 1, 5 ) ).toBe( "alpha" );
				expect( s.getCellType( wb, 1, 5 ) ).toBe( "string" );
			});
		});

		it( "Values in query data fall back to the query column type if they don't match the overridden type", function(){
			var data = QueryNew( "Number,String,Date,Time,Boolean", "VarChar,VarChar,VarChar,VarChar,VarChar", [ [ "01234", "alpha", "alpha", "alpha" , "alpha"] ] );
			var datatypes = { numeric: [ 1, 2 ], date: [ 3 ], time: [ 4 ], boolean: [ 5 ] };
			workbooks.Each( function( wb ){
				s.addRows( workbook=wb, data=data, datatypes=datatypes );
				expect( s.getCellValue( wb, 1, 1 ) ).toBe( 1234 );
				expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
				expect( s.getCellValue( wb, 1, 2 ) ).toBe( "alpha" );
				expect( s.getCellType( wb, 1, 2 ) ).toBe( "string" );
				expect( s.getCellValue( wb, 1, 3 ) ).toBe( "alpha" );
				expect( s.getCellType( wb, 1, 3 ) ).toBe( "string" );
				expect( s.getCellValue( wb, 1, 4 ) ).toBe( "alpha" );
				expect( s.getCellType( wb, 1, 4 ) ).toBe( "string" );
				expect( s.getCellValue( wb, 1, 5 ) ).toBe( "alpha" );
				expect( s.getCellType( wb, 1, 5 ) ).toBe( "string" );
			});
		});

		it( "Query data values with NO type override, default to query column types", function(){
			var data = QueryNew( "Number,String", "VarChar,VarChar", [ [ 1234, "01234" ] ] );
			var datatypes = { numeric: [ 2 ] };
			workbooks.Each( function( wb ){
				s.addRows( workbook=wb, data=data, datatypes=datatypes );
				expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
				expect( s.getCellType( wb, 1, 2 ) ).toBe( "numeric" );
			});
		});

		it( "Values in query data fall back to the autodetected type if they don't match the overridden type and ignoreQueryColumnDataTypes is true", function(){
			var data = QueryNew( "Number,String,Date,Time,Boolean", "VarChar,VarChar,VarChar,VarChar,VarChar", [ [ "01234", "alpha", "alpha", "alpha" , "alpha"] ] );
			var datatypes = { numeric: [ 1, 2 ], date: [ 3 ], time: [ 4 ], boolean: [ 5 ] };
			workbooks.Each( function( wb ){
				s.addRows( workbook=wb, data=data, ignoreQueryColumnDataTypes=true, datatypes=datatypes );
				expect( s.getCellValue( wb, 1, 1 ) ).toBe( 1234 );
				expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
				expect( s.getCellValue( wb, 1, 2 ) ).toBe( "alpha" );
				expect( s.getCellType( wb, 1, 2 ) ).toBe( "string" );
				expect( s.getCellValue( wb, 1, 3 ) ).toBe( "alpha" );
				expect( s.getCellType( wb, 1, 3 ) ).toBe( "string" );
				expect( s.getCellValue( wb, 1, 4 ) ).toBe( "alpha" );
				expect( s.getCellType( wb, 1, 4 ) ).toBe( "string" );
				expect( s.getCellValue( wb, 1, 5 ) ).toBe( "alpha" );
				expect( s.getCellType( wb, 1, 5 ) ).toBe( "string" );
			});
		});

		it( "Query data values in columns with an override type of 'auto' will have their type auto-detected, regardless of the query column type", function(){
			var data = QueryNew( "One,Two", "VarChar,VarChar", [ [ "2020-08-24", "2020-08-24" ], [ "3.1", "3.1" ] ] );
			var datatypes = { auto: [ 1 ] };
			workbooks.Each( function( wb ){
				s.addRows( workbook=wb, data=data, datatypes=datatypes );
				expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
				expect( s.getCellType( wb, 1, 2 ) ).toBe( "string" );
				expect( s.getCellType( wb, 2, 1 ) ).toBe( "numeric" );
				expect( s.getCellType( wb, 2, 2 ) ).toBe( "string" );
			});
		});

	});

	describe( "addRows throws an exception if", function(){

		it(
			title="adding more than 65536 rows to a binary spreadsheet",
			body=function(){
				var xls = workbooks[ 1 ];
				expect( function(){
					var rows = [];
					for( var i=1; i <= 65537; i++ ){
						rows.append( [ i ] );
					}
					var data = QueryNew( "ID","Integer",rows );
					variables.s.addRows( xls, data );
				}).toThrow( regex="Too many rows" );
			},
			skip=function(){
				return s.getIsACF();
			}
		);

		it( "the data is neither a query nor an array", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.addRows( wb, "string,list" );
				}).toThrow( regex="Invalid data" );
			});
		});

		it( "the data is an array which does not contain an array for each row", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.addRows( wb, [ { col1: "a" }, { col2: "b" } ] );// array of structs
				}).toThrow( regex="Invalid data" );
			});
		});

	});

});	
</cfscript>