<cfscript>
describe( "addRow", function(){

	beforeEach( function(){
		variables.data = "a,b";
		variables.dataAsArray = [ "a", "b" ];
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Appends a row with the minimum arguments", function(){
		workbooks.Each( function( wb ){
			s.addRow( wb, data );
			s.addRow( wb, "c,d" );// should be inserted at row 2
			var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
			var actual = s.sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Appends a row including commas with a custom delimiter", function(){
		workbooks.Each( function( wb ){
			s.addRow( workbook=wb, data="a,b|c,d", delimiter="|" );
			var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a,b", "c,d" ] ] );
			var actual = s.sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Appends a row as an array with the minimum arguments", function(){
		workbooks.Each( function( wb ){
			s.addRow( wb, dataAsArray );
			s.addRow( wb, [ "c", "d" ] );// should be inserted at row 2
			var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
			var actual = s.sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Inserts a row at a specifed position", function(){
		var expected = QueryNew( "column1,column2,column3", "VarChar,VarChar,VarChar", [ [ "a", "b", "" ], [ "c", "d", "" ], [ "", "e", "f" ] ] );
		workbooks.Each( function( wb ){
			s.addRow( wb, data );
			s.addRow( wb, "e,f", 2, 2 );
			s.addRow( wb, "c,d", 2, 1 );
			var actual = s.sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		});
		//array data
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			s.addRow( wb, dataAsArray );
			s.addRow( wb, [ "e", "f" ], 2, 2 );
			s.addRow( wb, [ "c", "d" ], 2, 1 );
			var actual = s.sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		});
	});

	it( "Replaces a row if insert is false", function(){
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ] ] );
		workbooks.Each( function( wb ){
			s.addRow( wb, data );
			s.addRow( workbook=wb, data=data, row=1, insert=false );
			var actual = s.sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			//array data
			s.addRow( wb, dataAsArray );
			s.addRow( workbook=wb, data=dataAsArray, row=1, insert=false );
			var actual = s.sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Handles embedded commas in comma delimited list data", function(){
		workbooks.Each( function( wb ){
			s.addRow( workbook=wb, data="'a,b', 'c,d'" );
			var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a,b", "c,d" ] ] );
			var actual = s.sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Adds numeric values correctly", function(){
		workbooks.Each( function( wb ){
			var data = "1,1.1";
			s.addRow( wb, data );
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( 1 );
			expect( s.getCellValue( wb, 1, 2 ) ).toBe( 1.1 );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 1, 2 ) ).toBe( "numeric" );
		});
			//array data
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			var dataAsArray = [ 1, 1.1 ];
			s.addRow( wb, dataAsArray );
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( 1 );
			expect( s.getCellValue( wb, 1, 2 ) ).toBe( 1.1 );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 1, 2 ) ).toBe( "numeric" );
		});
	});

	it( "Adds boolean values as strings", function(){
		workbooks.Each( function( wb ){
			var data = true;
			s.addRow( wb, data );
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( true );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
		});
		//array data
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			var dataAsArray = [ true ];
			s.addRow( wb, dataAsArray );
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( true );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
		});
	});

	it( "Adds date/time values correctly", function(){
		workbooks.Each( function( wb ){
			var dateValue = CreateDate( 2015, 04, 12 );
			var timeValue = CreateTime( 1, 0, 0 );
			var dateTimeValue = CreateDateTime( 2015, 04, 12, 1, 0, 0 );
			var data = "#dateValue#,#timeValue#,#dateTimeValue#";
			s.addRow( wb, data );
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( dateValue );
			expect( s.getCellValue( wb, 1, 2 ) ).toBe( timeValue );
			expect( s.getCellValue( wb, 1, 3 ) ).toBe( dateTimeValue );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 1, 2 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 1, 3 ) ).toBe( "numeric" );
		});
		//array data
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			var dateValue = CreateDate( 2015, 04, 12 );
			var timeValue = CreateTime( 1, 0, 0 );
			var dateTimeValue = CreateDateTime( 2015, 04, 12, 1, 0, 0 );
			var dataAsArray = [ dateValue, timeValue, dateTimeValue ];
			s.addRow( wb, dataAsArray );
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( dateValue );
			expect( s.getCellValue( wb, 1, 2 ) ).toBe( timeValue );
			expect( s.getCellValue( wb, 1, 3 ) ).toBe( dateTimeValue );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 1, 2 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 1, 3 ) ).toBe( "numeric" );
		});
	});

	it( "Adds zeros as zeros, not booleans", function(){
		workbooks.Each( function( wb ){
			s.addRow( wb, 0 );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
		});
		//array data
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			s.addRow( wb, [ 0 ] );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
		});
	});

	it( "Adds strings with leading zeros as strings not numbers", function(){
		workbooks.Each( function( wb ){
			s.addRow( wb, "01" );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
		});
		//array data
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			s.addRow( wb, [ "01" ] );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
		});
	});

	it( "Can insert more than 4009 rows containing dates without triggering an exception", function(){
		workbooks.Each( function( wb ){
			for( var i=1; i LTE 4010; i++ ){
				variables.s.addRow( wb, "2016-07-14" );
			}
		});
	});

	it( "Doesn't error if the workbook is SXSSF and autoSizeColumns is true", function(){
		var wb = s.newStreamingXlsx();
		s.addRow( workbook=local.wb, data=data, autoSizeColumns=true );
	});

	describe( "addRow() data type overriding",function(){

		it( "throws an error if invalid types are specified in the datatype struct", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					var data = [ "a", "b" ];
					var datatypes = { numeric: [ 1 ], varchar: [ 2 ] };
					s.addRow( workbook=wb, data=data, datatypes=datatypes );
				}).toThrow( regex="Invalid datatype\(s\)" );
			});
		});

		it( "throws an error if columns to override are not specified as arrays in the datatype struct", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					var data = [ "a", "b" ];
					var datatypes = { numeric: "1", string: "2" };
					s.addRow( workbook=wb, data=data, datatypes=datatypes );
				}).toThrow( regex="Invalid datatype\(s\)" );
			});
		});

		it( "Allows column data types to be overridden", function(){
			workbooks.Each( function( wb ){
				var datatypes = { numeric: [ 1 ], string: [ 2 ] };// can't test dates: date strings are always converted correctly!
				var data = "01234,1234567890123456";
				s.addRow( wb, data );
				expect( s.getCellValue( wb, 1, 1 ) ).toBe( "01234" );
				expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
				expect( s.getCellType( wb, 1, 2 ) ).toBe( "numeric" );
				s.addRow( workbook=wb, data=data, datatypes=datatypes );
				expect( s.getCellValue( wb, 2, 1 ) ).toBe( "1234" );
				expect( s.getCellType( wb, 2, 1 ) ).toBe( "numeric" );
				expect( s.getCellType( wb, 2, 2 ) ).toBe( "string" );
				// array data
				data = [ "01234", 1234567890123456 ];
				s.addRow( workbook=wb, data=data, datatypes=datatypes );
				expect( s.getCellValue( wb, 3, 1 ) ).toBe( "1234" );
				expect( s.getCellType( wb, 3, 1 ) ).toBe( "numeric" );
				expect( s.getCellType( wb, 3, 2 ) ).toBe( "string" );
			});
		});

		it( "Values fall back to the autodetected type if they don't match the overridden type", function(){
			workbooks.Each( function( wb ){
				var datatypes = { numeric: [ 1, 2 ] };
				var data = "01234,alpha";
				s.addRow( workbook=wb, data=data, datatypes=datatypes );
				expect( s.getCellValue( wb, 1, 1 ) ).toBe( 1234 );
				expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
				expect( s.getCellValue( wb, 1, 2 ) ).toBe( "alpha" );
				expect( s.getCellType( wb, 1, 2 ) ).toBe( "string" );
				data = [ "01234", "alpha" ];
				s.addRow( workbook=wb, data=data, datatypes=datatypes );
				expect( s.getCellValue( wb, 2, 1 ) ).toBe( 1234 );
				expect( s.getCellType( wb, 2, 1 ) ).toBe( "numeric" );
				expect( s.getCellValue( wb, 2, 2 ) ).toBe( "alpha" );
				expect( s.getCellType( wb, 2, 2 ) ).toBe( "string" );
			});
		});

	});

	describe( "addRow throws an exception if", function(){

		it( "row is zero or less", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.addRow( workbook=wb, data=data, row=0 );
				}).toThrow( regex="Invalid row" );
			});
		});

		it( "column is zero or less", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.addRow( workbook=wb, data=data, column=0 );
				}).toThrow( regex="Invalid column" );
			});
		});

		it( "insert is false and no row specified", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.addRow( workbook=wb, data=data, insert=false );
				}).toThrow( regex="Missing row" );
			});
		});

	});

});	
</cfscript>