<cfscript>
describe( "addColumn", function(){

	beforeEach( function(){
		variables.columnData = "a,b";
		variables.dataAsArray = [ "a", "b" ];
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Adds a column with the minimum arguments", function(){
		workbooks.Each( function( wb ){
			s.addColumn( wb, columnData );
			var expected = QueryNew( "column1", "VarChar", [ [ "a" ], [ "b" ] ] );
			var actual = s.sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Adds a column with the minimum arguments using array data", function(){
		workbooks.Each( function( wb ){
			s.addColumn( wb, dataAsArray );
			var expected = QueryNew( "column1", "VarChar", [ [ "a" ], [ "b" ] ] );
			var actual = s.sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Adds a column at a given start row", function(){
		workbooks.Each( function( wb ){
			s.addColumn( wb, columnData, 2 );
			var expected = QueryNew( "column1", "VarChar", [ [ "" ], [ "a" ], [ "b" ] ] );
			var actual = s.sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		});
	});

	it( "Adds a column at a given column number", function(){
		workbooks.Each( function( wb ){
			s.addColumn( workbook=wb, data=columnData, startColumn=2 );
			var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "a" ], [ "", "b" ] ] );
			var actual = s.sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		});
	});

	it( "Adds a column including commas with a custom delimiter", function(){
		workbooks.Each( function( wb ){
			var columnData = "a,b|c,d";
			s.addColumn( workbook=wb, data=columnData,delimiter="|" );
			var expected = QueryNew( "column1", "VarChar", [ [ "a,b" ], [ "c,d" ] ] );
			var actual = s.sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Inserts (not replaces) a column with the minimum arguments", function(){
		workbooks.Each( function( wb ){
			s.addColumn( wb, columnData );
			s.addColumn( workbook=wb, data=columnData, insert=true );
			var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a","a" ], [ "b","b" ] ] );
			var actual = s.sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Adds numeric values correctly", function(){
		workbooks.Each( function( wb ){
			var rowData = "1,1.1";
			s.addColumn( wb, rowData );
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( 1 );
			expect( s.getCellValue( wb, 2, 1 ) ).toBe( 1.1 );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 2, 1 ) ).toBe( "numeric" );
		});
	});

	it( "Adds boolean values as strings", function(){
		workbooks.Each( function( wb ){
			var rowData = true;
			s.addColumn( wb, rowData );
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( true );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
		});
	});

	it( "Adds date/time values correctly", function(){
		workbooks.Each( function( wb ){
			var dateValue = CreateDate( 2015, 04, 12 );
			var timeValue = CreateTime( 1, 0, 0 );
			var dateTimeValue = CreateDateTime( 2015, 04, 12, 1, 0, 0 );
			var rowData = "#dateValue#,#timeValue#,#dateTimeValue#";
			s.addColumn( wb, rowData );
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( dateValue );
			expect( s.getCellValue( wb, 2, 1 ) ).toBe( timeValue );
			expect( s.getCellValue( wb, 3, 1 ) ).toBe( dateTimeValue );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 2, 1 ) ).toBe( "numeric" );
			expect( s.getCellType( wb, 3, 1 ) ).toBe( "numeric" );
		});
	});

	it( "Adds zeros as zeros, not booleans", function(){
		workbooks.Each( function( wb ){
			s.addColumn( wb, 0 );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
		});
	});

	it( "Adds strings with leading zeros as strings not numbers", function(){
		workbooks.Each( function( wb ){
			s.addColumn( wb, "01" );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
		});
	});

});	
</cfscript>