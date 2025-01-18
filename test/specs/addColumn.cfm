<cfscript>
describe( "addColumn", function(){

	beforeEach( function(){
		variables.columnData = "a,b";
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Adds a column with the minimum arguments", function(){
		workbooks.Each( function( wb ){
			s.addColumn( wb, columnData );
			var expected = QueryNew( "column1", "VarChar", [ [ "a" ], [ "b" ] ] );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Adds a column with the minimum arguments using array data", function(){
		workbooks.Each( function( wb ){
			s.addColumn( wb, columnData.ListToArray() );
			var expected = QueryNew( "column1", "VarChar", [ [ "a" ], [ "b" ] ] );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Adds a column at a given NEW start row", function(){
		workbooks.Each( function( wb ){
			s.addColumn( workbook=wb, data=columnData, startRow=2 );
			var expected = QueryNew( "column1", "VarChar", [ [ "" ], [ "a" ], [ "b" ] ] );
			var actual = s.getSheetHelper().sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		});
	});

	it( "Adds a column at a given EXISTING start row", function(){
		workbooks.Each( function( wb ){
			s.addRows( wb, [ [ "x" ], [ "y" ] ] )
				.addColumn( workbook=wb, data=columnData, startRow=2 );
			var expected = querySim( "column1,column2
				x|
				y|a
				 |b
			");
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Adds a column at a given column number", function(){
		workbooks.Each( function( wb ){
			s.addColumn( workbook=wb, data=columnData, startColumn=2 );
			var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "a" ], [ "", "b" ] ] );
			var actual = s.getSheetHelper().sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		});
	});

	it( "Adds a column including commas with a custom delimiter", function(){
		workbooks.Each( function( wb ){
			var columnData = "a,b|c,d";
			s.addColumn( workbook=wb, data=columnData, delimiter="|" );
			var expected = QueryNew( "column1", "VarChar", [ [ "a,b" ], [ "c,d" ] ] );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "Allows the data type to be specified", function(){
		var columnData = [ 1.234 ];
		workbooks.Each( function( wb ){
			s.addColumn( workbook=wb, data=columnData, datatype="string" );
			var actual = s.getCellValue( wb, 1, 1 );
			expect( actual ).toBe( "1.234" );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
		});
	});
		

	describe( "Insert behaviour", function(){

		it( "Inserts column after existing columns by default", function(){
			workbooks.Each( function( wb ){
				s.addColumn( wb, columnData )
					.addColumn( wb, [ "c", "d" ] );
				var expected = querySim( "column1,column2
					a|c
					b|d
				");
				var actual = s.getSheetHelper().sheetToQuery( wb );
				expect( actual ).toBe( expected );
			});
		});

		it( "By default, overwrites an existing column if 'startColumn' is specified", function(){
			workbooks.Each( function( wb ){
				s.addColumn( wb, "a,b" )
					.addColumn( workbook=wb, data="x,y", startColumn=1 );
				var expected = QueryNew( "column1", "VarChar", [ [ "x" ], [ "y" ] ] );
				var actual = s.getSheetHelper().sheetToQuery( wb );
				expect( actual ).toBe( expected );
				s.addColumn( wb, [ "a", "b" ] )
					.addColumn( workbook=wb, data=columnData, startColumn=2 );
				var expected = querySim( "column1,column2
					x|a
					y|b
				");
				var actual = s.getSheetHelper().sheetToQuery( wb );
				expect( actual ).toBe( expected );
			});
		});

		it( "Shifts columns to the right if startColumn is specified and column already exists and 'insert=true'", function(){
			workbooks.Each( function( wb ){
				s.addColumn( wb, [ "a", "b" ] )
					.addColumn( wb, [ "c", "d" ] )
					.addColumn( wb, [ "e", "f" ] )
					.addColumn( workbook=wb, data="x,y", startColumn=2, insert=true );
				var expected = querySim( "column1,column2,column3,column4
					a|x|c|e
					b|y|d|f
				");
				var actual = s.getSheetHelper().sheetToQuery( wb );
				expect( actual ).toBe( expected );
			});
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
			var timeValue = _CreateTime( 1, 0, 0 );
			var dateTimeValue = CreateDateTime( 2015, 04, 12, 1, 0, 0 );
			var rowData = "#dateValue#,#timeValue#,#dateTimeValue#";
			s.addColumn( wb, rowData );
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( "2015-04-12" );
			expect( s.getCellValue( wb, 2, 1 ) ).toBe( "01:00:00" );
			expect( s.getCellValue( wb, 3, 1 ) ).toBe( "2015-04-12 01:00:00" );
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

	it( "Is chainable", function(){
		workbooks.Each( function( wb ){
			s.newChainable( wb ).addColumn( columnData );
			var expected = QueryNew( "column1", "VarChar", [ [ "a" ], [ "b" ] ] );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

});	
</cfscript>