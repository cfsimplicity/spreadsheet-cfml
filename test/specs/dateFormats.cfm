<cfscript>
describe( "dateFormats customisability",function(){

	it( "the default dateFormats can be overridden individually",function() {
		local.s = newSpreadsheetInstance();
		var expected = {
			DATE: "yyyy-mm-dd"
			,DATETIME: "yyyy-mm-dd HH:nn:ss"
			,TIME: "hh:mm:ss"
			,TIMESTAMP: "yyyy-mm-dd hh:mm:ss"
		};
		var actual = s.getDateFormats();
		expect( actual ).toBe( expected );
		local.s = newSpreadsheetInstance( dateFormats={ DATE="mm/dd/yyyy" } );
		expected = {
			DATE: "mm/dd/yyyy"
			,DATETIME: "yyyy-mm-dd HH:nn:ss"
			,TIME: "hh:mm:ss"
			,TIMESTAMP: "yyyy-mm-dd hh:mm:ss"
		};
		actual = s.getDateFormats();
		expect( actual ).toBe( expected );
	});

	it( "allows the format of date and time values to be customised", function() {
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		//Dates
		var dateValue =  CreateDate( 2019, 04, 12 );
		var timeValue = CreateTime( 1, 5, 5 );
		var timestampValue = CreateDateTime(  2019, 04, 12, 1, 5, 5 );
		workbooks.Each( function( wb ){
			s.setCellValue( wb, dateValue, 1, 1 );
			var expected = DateFormat( dateValue, "yyyy-mm-dd" );
			var actual = s.getCellValue( wb, 1, 1 );
			expect( actual ).toBe( expected );
			//Times
			s.setCellValue( wb, timeValue, 1, 1 );
			expected = TimeFormat( timeValue, "hh:mm:ss" );
			actual = s.getCellValue( wb, 1, 1 );
			expect( actual ).toBe( expected );
			//timestamps
			s.setCellValue( wb, timestampValue, 1, 1 );
			expected = DateTimeFormat( timestampValue, "yyyy-mm-dd hh:nn:ss" );
			actual = s.getCellValue( wb, 1, 1 );
			expect( actual ).toBe( expected );

			// custom date format
			local.s = newSpreadsheetInstance( dateFormats={ DATE="mm/dd/yyyy" } );
			s.setCellValue( wb, dateValue, 1, 1 );
			expected = DateFormat( dateValue, "mm/dd/yyyy" );
			actual = s.getCellValue( wb, 1, 1 );
			expect( actual ).toBe( expected );
			//custom time format
			local.s = newSpreadsheetInstance( dateFormats={ TIME="h:m:s" } );
			s.setCellValue( wb, timeValue, 1, 1 );
			expected = TimeFormat( timeValue, "h:m:s" );
			actual = s.getCellValue( wb, 1, 1 );
			//custom timestamp format
			local.s = newSpreadsheetInstance( dateFormats={ TIMESTAMP="mm/dd/yyyy h:m:s" } );
			s.setCellValue( wb, timestampValue, 1, 1 );
			expected = DateTimeFormat( timestampValue, "mm/dd/yyyy h:n:s" );
			actual = s.getCellValue( wb, 1, 1 );
		});
	});

	it( "Uses the overridden DATETIME format mask when generating CSV and HTML",function() {
		local.s = newSpreadsheetInstance( dateFormats={ DATETIME="mm/dd/yyyy h:n:s" } );
		var path = getTestFilePath( "test.xls" );
		var actual = s.read( src=path, format="html" );
		var expected = "<tbody><tr><td>a</td><td>b</td></tr><tr><td>1</td><td>04/01/2015 12:0:0</td></tr><tr><td>04/01/2015 1:1:1</td><td>2</td></tr></tbody>";
		expect( actual ).toBe( expected );
		expected = 'a,b#crlf#1,04/01/2015 12:0:0#crlf#04/01/2015 1:1:1,2';
		actual = s.read( src=path, format="csv" );
		expect( actual ).toBe( expected );
	});

	describe( "dateFormats: throws an exception if",function(){

		it( "a passed format key is invalid",function() {
			expect( function(){
				local.s = newSpreadsheetInstance( dateFormats={ DAT="mm/dd/yyyy" } );
			}).toThrow( regex="Invalid date format key" );
		});

	});	

});	
</cfscript>