<cfscript>
describe( "cellValue", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Gets the value from the specified cell", function(){
		var data =  [ [ "a", "b" ], [ "c", "d" ] ];
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

	it( "returns the visible/formatted value by default", function(){
		var value = 0.000011;
		workbooks.Each( function( wb ){
			s.setCellValue( wb, value, 1, 1 );
			s.formatCell( wb, { dataformat: "0.00000" }, 1, 1 );
			var actual = s.getCellValue( wb, 1, 1 );
			expect( actual ).toBe( 0.00001 );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
			var decimalHasBeenOutputInScientificNotation = ( Trim( actual ).FindNoCase( "E" ) > 0 );
			expect( decimalHasBeenOutputInScientificNotation ).toBeFalse();
		});
	});

	it( "can return the raw (unformatted) value", function(){
		var value = 0.000011;
		workbooks.Each( function( wb ){
			s.setCellValue( wb, value, 1, 1 );
			s.formatCell( wb, { dataformat: "0.00000" }, 1, 1 );
			var actual = s.getCellValue( wb, 1, 1, false );
			expect( actual ).toBe( value );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "numeric" );
			// chainable
			var actual = s.newChainable( wb ).getCellValue( 1, 1, false );
			expect( actual ).toBe( value );
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

		it( "allows forcing values to be added as times (without a date)", function(){
			var value = "08:21:30";
			workbooks.Each( function( wb ){
				s.setCellValue( wb, value, 1, 1, "time" );
				var actual = s.getCellValue( wb, 1, 1 );
				expect( actual ).toBe( value );
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

		it( "support legacy 'type' argument name", function(){
			var value = 1.234;
			workbooks.Each( function( wb ){
				s.setCellValue( workbook=wb, value=value, row=1, column=1, type="string" );
				var actual = s.getCellValue( wb, 1, 1 );
				expect( actual ).toBe( "1.234" );
				expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
			});
			workbooks.Each( function( wb ){
				var actual = s.newChainable( wb )
				.setCellValue( value=value, row=1, column=1, type="string" )
				.getCellValue( 1, 1 );
			expect( actual ).toBe( "1.234" );
			expect( s.getCellType( wb, 1, 1 ) ).toBe( "string" );
		});
		});

	});

	describe(
		title="Lucee only timezone tests",
		body=function(){

			it( "Sets the specified cell to the specified date value even if the Lucee timezone doesn't match the system", function(){
				variables.currentTZ = GetTimeZone();
				//Needs manually adjusting if the test Lucee instance TZ is in Central European Time, i.e. same as London e.g. Lisbon
				variables.tempTZ = ( currentTZ == "Europe/London" )? "Europe/Paris": "Europe/London";
				SetTimeZone( tempTZ );
				var value = CreateDate( 2015, 04, 12 );
				workbooks.Each( function( wb ){
					s.setCellValue( wb, value, 1, 1 );
					s.formatCell( wb, { dataformat: "0.0" }, 1, 1 );
					expect( s.getCellValue( wb, 1, 1 ) ).toBe( 42106.0 );// whole number = date, no time
				});
				SetTimeZone( currentTZ );
			});

		},
		skip=function(){
			return s.getIsACF();
		}
	);

	describe( "setCellValue throws an exception if", function(){

		it( "the data type is invalid", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.setCellValue( wb, "test", 1, 1, "blah" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidDatatype" );
			});
		});

	});

});	
</cfscript>