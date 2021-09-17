<cfscript>
describe( "cellValue", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Gets the value from the specified cell", function(){
		var data =  QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a","b" ], [ "c","d" ] ] );
		workbooks.Each( function( wb ){
			s.addRows( wb, data );
			expect( s.getCellValue( wb, 2, 2 ) ).toBe( "d" );
		});
	});

	it( "Sets the specified cell to the specified string value", function(){
		var value = "test";
		workbooks.Each( function( wb ){
			s.setCellValue( wb, value, 1, 1 );
			var actual = s.getCellValue( wb, 1, 1 );
			expect( actual ).toBe( value );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
		});
	});

	it( "getCellValue and setCellValue are chainable", function(){
		var value = "test";
		workbooks.Each( function( wb ){
			var actual = s.newChainable( wb )
				.setCellValue(value, 1, 1 )
				.getCellValue( 1, 1 );
			expect( actual ).toBe( value );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
		});
	});

	it( "Sets the specified cell to the specified numeric value", function(){
		var value = 1;
		workbooks.Each( function( wb ){
			s.setCellValue( wb, value, 1, 1 );
			var actual = s.getCellValue( wb, 1, 1 );
			expect( actual ).toBe( value );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
		});
	});

	it( "Sets the specified cell to the specified date value", function(){
		var value = CreateDate( 2015, 04, 12 );
		workbooks.Each( function( wb ){
			s.setCellValue( wb, value, 1, 1 );
			var expected = DateFormat( value, "yyyy-mm-dd" );
			var actual = s.getCellValue( wb, 1, 1 );
			expect( actual ).toBe( expected );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
		});
	});

	it( "Sets the specified cell to the specified boolean value with a data type of string by default", function(){
		var value = true;
		workbooks.Each( function( wb ){
			s.setCellValue( wb, value, 1, 1 );
			var actual = s.getCellValue( wb, 1, 1 );
			expect( actual ).toBe( value );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
		});
	});

	it( "Sets the specified range of cells to the specified value", function(){
		var value = "a";
		var expected = querySim(
				"column1,column2
				a|a
				a|a");
		workbooks.Each( function( wb ){
			s.setCellRangeValue( wb, value, 1, 2, 1, 2 );
			actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "handles numbers with leading zeros correctly", function(){
		var value = "0162220494";
		workbooks.Each( function( wb ){
		s.setCellValue( wb, value, 1, 1 );
			var actual = s.getCellValue( wb, 1, 1 );
			expect( actual ).toBe( value );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
		});
	});

	it( "handles non-date values correctly that Lucee parses as partial dates far in the future", function(){
		workbooks.Each( function( wb ){
			var value = "01-23112";
			s.setCellValue( wb, value, 1, 1 );
			var actual = s.getCellValue( wb, 1, 1 );
			expect( actual ).toBe( value );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
			value = "23112-01";
			s.setCellValue( wb, value, 1, 1 );
			actual = s.getCellValue( wb, 1, 1 );
			expect( actual ).toBe( value );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
		});
	});

	it( "does not accept '9a' or '9p' or '9 a' as valid dates, correcting ACF", function(){
		values = [ "9a", "9p", "9 a", "9    p", "9A" ];
		values.Each( function( value ){
			workbooks.Each( function( wb ){
				s.setCellValue( wb, value, 1, 1 );
				var actual = s.getCellValue( wb, 1, 1 );
				expect( actual ).toBe( value );
				expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
			});
		});
	});

	it( "but does accept date strings with AM or PM", function(){
		workbooks.Each( function( wb ){
			s.setCellValue( wb, "8/22/2020 10:34 AM", 1, 1 );
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( "2020-08-22 10:34:00" );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
			s.setCellValue( wb, "12:53 pm", 1, 1 );
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( "12:53:00" );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
		});
	});

	describe( "allows the auto data type detection to be overridden", function(){

		it( "allows forcing values to be added as strings", function(){
			var value = 1.234;
			workbooks.Each( function( wb ){
				s.setCellValue( wb, value, 1, 1, "string" );
				var actual = s.getCellValue( wb, 1, 1 );
				expect( actual ).toBe( "1.234" );
				expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
			});
		});

		it( "allows forcing values to be added as numbers", function(){
			var value = "0123";
			workbooks.Each( function( wb ){
				s.setCellValue( wb, value, 1, 1, "numeric" );
				var actual = s.getCellValue( wb, 1, 1 );
				expect( actual ).toBe( 123 );
				expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
			});
		});

		it( "allows forcing values to be added as dates", function(){
			var value = "01.1990";
			workbooks.Each( function( wb ){
				s.setCellValue( wb, value, 1, 1, "date" );
				var actual = s.getCellValue( wb, 1, 1 );
				expect( DateFormat( actual, "yyyy-mm-dd" ) ).toBe( "1990-01-01" );
				expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );// dates are numeric in Excel
			});
		});

		it( "allows forcing values to be added as booleans", function(){
			var values = [ "true", true, 1, "1", "yes", 10 ];
			workbooks.Each( function( wb ){
				for( var value in values ){
					s.setCellValue( wb, value, 1, 1, "boolean" );
					var actual = s.getCellValue( wb, 1, 1 );
					expect( actual ).toBeTrue();
					expect( s.getCellType( wb, 1, 1 ) ).toBe( "boolean" );
				}
			});
		});

		it( "allows forcing values to be added as blanks", function(){
			var values = [ "", "blah" ];
			workbooks.Each( function( wb ){
				for( var value in values ){
					s.setCellValue( wb, value, 1, 1, "blank" );
					var actual = s.getCellValue( wb, 1, 1 );
					expect( actual ).toBeEmpty();
					expect( s.getCellType( wb, 1, 1 ) ).toBe( "blank" );	
				}
			});
		});

	});

	describe( "setCellValue throws an exception if", function(){

		it( "the data type is invalid", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.setCellValue( wb, "test", 1, 1, "blah" );
				}).toThrow( regex="Invalid data type" );
			});
		});

	});

});	
</cfscript>