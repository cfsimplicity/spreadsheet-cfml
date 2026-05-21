<cfscript>
describe( "dateFormats customisability", ()=>{
	
	it( "the default dateFormats can be overridden individually on init", ()=>{
		// Default formats loaded
		var defaultFormats = s.getDateHelper().defaultFormats();
		local.s = newSpreadsheetInstance();
		var expected = defaultFormats
		var actual = local.s.getDateFormats();
		expect( actual ).toBe( expected );
		// Override date mask pre instance creation
		local.s = newSpreadsheetInstance( dateFormats={ DATE: "mm/dd/yyyy" } );
		expected.DATE = "mm/dd/yyyy";
		actual = local.s.getDateFormats();
		expect( actual ).toBe( expected );
	})

	it( "the dateFormats can be set post-init", ()=>{
		// Default formats loaded
		var defaultFormats = s.getDateHelper().defaultFormats();
		local.s = newSpreadsheetInstance();
		var expected = defaultFormats
		var actual = local.s.getDateFormats();
		expect( actual ).toBe( expected );

		// Override date mask post instance creation
		var customDateFormats = { DATE: "mm/dd/yyyy" };
		local.s.setDateFormats( customDateFormats );
		expected.DATE = customDateFormats.DATE;
		expect( local.s.getDateFormats() ).toBe( expected );
	})

	it( "allows the format of date and time values to be customised", ()=>{
		// Formats change between engines
		var defaultFormats = s.getDateHelper().defaultFormats();
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		//Dates
		var dateValue =  CreateDate( 2019, 04, 12 );
		var timeValue = _CreateTime( 1, 5, 5 );
		var timestampValue = CreateDateTime(  2019, 04, 12, 1, 5, 5 );
		workbooks.Each( ( wb )=>{
			s.setCellValue( wb, dateValue, 1, 1 );
			var expected = DateFormat( dateValue, defaultFormats.DATE );
			var actual = s.getCellValue( wb, 1, 1 );
			expect( actual ).toBe( expected );
			//Times
			s.setCellValue( wb, timeValue, 1, 1 );
			expected = TimeFormat( timeValue, defaultFormats.TIME );
			actual = s.getCellValue( wb, 1, 1 );
			expect( actual ).toBe( expected );
			//timestamps
			s.setCellValue( wb, timestampValue, 1, 1 );
			expected = DateTimeFormat( timestampValue, defaultFormats.DATETIME );
			actual = s.getCellValue( wb, 1, 1 );
			expect( actual ).toBe( expected );

			// Custom format (changes between engines)
			var customDateFormat = !s.getIsBoxlang() ? "mm/dd/yyyy" : "MM/dd/yyyy";
			local.s = newSpreadsheetInstance( dateFormats={ DATE=customDateFormat } );
			s.setCellValue( wb, dateValue, 1, 1 );
			expected = DateFormat( dateValue, customDateFormat );
			actual = s.getCellValue( wb, 1, 1 );
			expect( actual ).toBe( expected );
			//custom time format
			local.s = newSpreadsheetInstance( dateFormats={ TIME="h:m:s" } );
			s.setCellValue( wb, timeValue, 1, 1 );
			expected = TimeFormat( timeValue, "h:m:s" );
			actual = s.getCellValue( wb, 1, 1 );
			//custom timestamp format (changes between engines)
			var customTimestampFormat = !s.getIsBoxlang() ? "mm/dd/yyyy h:m:s" : "MM/dd/yyyy h:m:s"
			local.s = newSpreadsheetInstance( dateFormats={ TIMESTAMP=customTimestampFormat } );
			s.setCellValue( wb, timestampValue, 1, 1 );
			expected = DateTimeFormat( timestampValue, customTimestampFormat );
			actual = s.getCellValue( wb, 1, 1 );
		})
	})

	it( "Uses the overridden DATETIME format mask when generating CSV and HTML",()=>{
		var customDateTimeFormat = !s.getIsBoxlang() ? "mm/dd/yyyy h:n:s" : "MM/dd/yyyy h:m:s";
		local.s = newSpreadsheetInstance( dateFormats={ DATETIME=customDateTimeFormat } );
		var path = getTestFilePath( "test.xls" );
		var actual = s.read( src=path, format="html" );
		var expected = "<tbody><tr><td>a</td><td>b</td></tr><tr><td>1</td><td>04/01/2015 12:0:0</td></tr><tr><td>04/01/2015 1:1:1</td><td>2</td></tr></tbody>";
		expect( actual ).toBe( expected );
		expected = 'a,b#newline#1,04/01/2015 12:0:0#newline#04/01/2015 1:1:1,2#newline#';
		actual = s.read( src=path, format="csv" );
		expect( actual ).toBe( expected );
	})

	describe( "dateFormats: throws an exception if",()=>{

		it( "a passed format key is invalid",()=>{
			expect( ()=>{
				local.s = newSpreadsheetInstance( dateFormats={ DAT="mm/dd/yyyy" } );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidDateFormatKey" );
		})

	})	

})	
</cfscript>