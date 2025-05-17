<cfscript>
describe( "read", ()=>{

	it( "Can read a traditional XLS file", ()=>{
		var path = getTestFilePath( "test.xls" );
		var workbook = s.read( src=path );
		expect( s.isBinaryFormat( workbook ) ).toBeTrue();
	})

	it( "Can read an OOXML file", ()=>{
		var path = getTestFilePath( "test.xlsx" );
		var workbook = s.read( src=path );
		expect( s.isXmlFormat( workbook ) ).toBeTrue();
	})

	it( "Is chainable", ()=>{
		var path = getTestFilePath( "test.xls" );
		var workbook = s.newChainable()
			.read( path )
			.getWorkbook();
		expect( s.isBinaryFormat( workbook ) ).toBeTrue();
		path = getTestFilePath( "test.xlsx" );
		workbook = s.newChainable()
			.read( path )
			.getWorkbook();
		expect( s.isXmlFormat( workbook ) ).toBeTrue();
	})

	it( "Ends the chain and returns the import result if format is specified", ()=>{
		variables.spreadsheetTypes.Each( ( type )=>{
			var path = getTestFilePath( "test.#type#" );
			var data = s.newChainable().read( path, "query" );
			expect( IsQuery( data ) ).toBeTrue();
		})
	})

	it( "Can return HTML table rows from an Excel file", ()=>{
		var path = getTestFilePath( "test.xls" );
		var actual = s.read( src=path, format="html" );
		var expected = "<tbody><tr><td>a</td><td>b</td></tr><tr><td>1</td><td>2015-04-01 00:00:00</td></tr><tr><td>2015-04-01 01:01:01</td><td>2</td></tr></tbody>";
		expect( actual ).toBe( expected );
		actual = s.read( src=path, format="html", headerRow=1 );
		expected = "<tbody><tr><td>1</td><td>2015-04-01 00:00:00</td></tr><tr><td>2015-04-01 01:01:01</td><td>2</td></tr></tbody>";
		expect( actual ).toBe( expected );
		actual = s.read( src=path, format="html", headerRow=1, includeHeaderRow=true );
		expected="<thead><tr><th>a</th><th>b</th></tr></thead><tbody><tr><td>a</td><td>b</td></tr><tr><td>1</td><td>2015-04-01 00:00:00</td></tr><tr><td>2015-04-01 01:01:01</td><td>2</td></tr></tbody>";
		expect( actual ).toBe( expected );
	})

	it( "Can return a CSV string from an Excel file", ()=>{
		var path = getTestFilePath( "test.xls" );
		var expected = 'a,b#newline#1,2015-04-01 00:00:00#newline#2015-04-01 01:01:01,2#newline#';
		var actual = s.read( src=path, format="csv" );
		expect( actual ).toBe( expected );
		expected = 'a,b#newline#a,b#newline#1,2015-04-01 00:00:00#newline#2015-04-01 01:01:01,2#newline#';
		actual = s.read( src=path, format="csv", headerRow=1, includeHeaderRow=true );
		expect( actual ).toBe( expected );
	})

	it( "Escapes double-quotes in string values when reading to CSV", ()=>{
		var data = QueryNew( "column1", "VarChar", [ [ 'a "so-called" test' ] ] );
		var workbook = s.new();
		s.addRows( workbook, data )
			.write( workbook, tempXlsPath, true );
		var expected = '"a ""so-called"" test"#newline#';
		var actual = s.read( src=tempXlsPath, format="csv" );
		expect( actual ).toBe( expected );
	})

	it( "Accepts a custom delimiter when generating CSV", ()=>{
		var path = getTestFilePath( "test.xls" );
		var expected = 'a|b#newline#1|2015-04-01 00:00:00#newline#2015-04-01 01:01:01|2#newline#';
		var actual = s.read( src=path, format="csv", csvDelimiter="|" );
		expect( actual ).toBe( expected );
	})

	it( "Includes columns formatted as 'hidden' by default", ()=>{
		spreadsheetTypes.Each( ( type )=>{
			var path = variables[ "temp" & type & "Path" ];
			s.newChainable( type )
				.addColumn( "a1" )
				.addColumn( "b1" )
				.hideColumn( 1 )
				.write( path, true );
			var actual = s.read( src=path, format="query" );
			var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a1", "b1" ] ] );
			expect( actual ).toBe( expected );
		})
	})

	it( "Can exclude columns formatted as 'hidden'", ()=>{
		var workbook = s.new();
		s.addColumn( workbook, "a1" )
			.addColumn( workbook, "b1" )
			.hideColumn( workbook, 1 )
			.write( workbook, tempXlsPath, true );
		var actual = s.read( src=tempXlsPath, format="query", includeHiddenColumns=false );
		var expected = QuerySim( "column2
			b1");
		expect( actual ).toBe( expected );
	})

	it( "Includes rows formatted as 'hidden' by default", ()=>{
		var data = QueryNew( "column1", "VarChar", [ [ "Apple" ], [ "Banana" ], [ "Carrot" ], [ "Doughnut" ] ] );
		spreadsheetTypes.Each( ( type )=>{
			var path = variables[ "temp" & type & "Path" ];
			s.newChainable( type )
				.addRows( data )
				.hideRow( 1 )
				.write( path, true );
			var actual = s.read( src=path, format="query" );
			var expected = data;
			expect( actual ).toBe( expected );
		})
	})

	it( "Can exclude rows formatted as 'hidden'", ()=>{
		var data = QueryNew( "column1", "VarChar", [ [ "Apple" ], [ "Banana" ], [ "Carrot" ], [ "Doughnut" ] ] );
		var workbook = s.new();
		s.addRows( workbook, data );
		s.hideRow( workbook, 1 );
		s.hideRow( workbook, 3 );
		s.write( workbook, tempXlsPath, true );
		var actual = s.read( src=tempXlsPath, format="query", includeHiddenRows=false );
		var expected = QueryNew( "column1", "VarChar", [ [ "Banana" ], [ "Doughnut" ] ] );
		expect( actual ).toBe( expected );
	})

	it( "Can read an encrypted XLSX file", ()=>{
		var path = getTestFilePath( "passworded.xlsx" );
		var expected = QueryNew( "column1", "VarChar", [ [ "secret" ] ] );
		var actual = s.read( src=path, format="query", password="pass" );
		expect( actual ).toBe( expected );
	})

	it( "Can read an encrypted binary file", ()=>{
		var path = getTestFilePath( "passworded.xls" );
		var expected = QueryNew( "column1", "VarChar", [ [ "secret" ] ] );
		var actual = s.read( src=path, format="query", password="pass" );
		expect( actual ).toBe( expected );
	})

	it( "Can read a spreadsheet containing a formula", ()=>{
		var workbook = s.new();
		s.addColumn( workbook, "1,1" );
		var theFormula = "SUM(A1:A2)";
		s.setCellFormula( workbook, theFormula, 3, 1 )
			.write( workbook=workbook, filepath=tempXlsPath, overwrite=true );
		var expected = QueryNew( "column1","Integer", [ [ 1 ], [ 1 ], [ 2 ] ] );
		var actual = s.read( src=tempXlsPath, format="query" );
		expect( actual ).toBe( expected );
	})

	describe( "read throws an exception if", ()=>{

		it( "queryColumnTypes is specified as a 'columnName/type' struct, but headerRow and columnNames arguments are missing", ()=>{
			expect( ()=>{
				var columnTypes = { col1: "Integer" };
				s.read( src=getTestFilePath( "test.xlsx" ), format="query", queryColumnTypes=columnTypes );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidQueryColumnTypesArgument" );
		})

		it( "the 'query' argument is passed", ()=>{
			expect( ()=>{
				s.read( src=tempXlsPath, query="q" );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidQueryArgument" );
		})

		it( "the format argument is invalid", ()=>{
			expect( ()=>{
				s.read( src=tempXlsPath, format="wrong" );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidReadFormat" );
		})

		it( "the file doesn't exist", ()=>{
			expect( ()=>{
				var path = getTestFilePath( "nonexistent.xls" );
				s.read( src=path );
			}).toThrow( type="cfsimplicity.spreadsheet.nonExistentFile" );
		})

		it( "the sheet name doesn't exist", ()=>{
			expect( ()=>{
				var path = getTestFilePath( "test.xls" );
				s.read( src=path, format="query", sheetName="nonexistent" );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetName" );
		})

		it( "the sheet number doesn't exist", ()=>{
			expect( ()=>{
				var path = getTestFilePath( "test.xls" );
				s.read( src=path, format="query", sheetNumber=20 );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetNumber" );
		})

		it( "both sheetName and sheetNumber arguments are specified", ()=>{
			expect( ()=>{
				var path = getTestFilePath( "test.xls" );
				s.read( src=path, sheetName="sheet1", sheetNumber=2 );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidArguments" );
		})

		it( "the password for an encrypted XML file is incorrect", ()=>{
			expect( ()=>{
				var tempXlsxPath = getTestFilePath( "passworded.xlsx" );
				s.read( src=tempXlsxPath, format="query", password="parse" );
			}).toThrow( regex="(Invalid password|Password incorrect|password is invalid)" );
		})

		it( "the password for an encrypted binary file is incorrect", ()=>{
			expect( ()=>{
				var xlsPath = getTestFilePath( "passworded.xls" );
				s.read( src=xlsPath, format="query", password="parse" );
			}).toThrow( regex="(Invalid password|Password incorrect|password is invalid)" );
		})

		it( "the source file is not a spreadsheet", ()=>{
			expect( ()=>{
				var path = getTestFilePath( "notaspreadsheet.txt" );
				s.read( src=path );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidFile" );
		})

		it( "the source file appears to contain CSV or TSV, and suggests using 'csvToQuery'", ()=>{
			expect( ()=>{
				var path = getTestFilePath( "csv.xls" );
				s.read( src=path );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidFile" );
			expect( ()=>{
				var path = getTestFilePath( "test.tsv" );
				s.read( src=path );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidFile" );
		})

		it( "the source file is in an old format not supported by POI", ()=>{
			expect( ()=>{
				var path = getTestFilePath( "oldformat.xls" );
				s.read( src=path );
			}).toThrow( type="cfsimplicity.spreadsheet.oldExcelFormatException" );
		})

	})

})
</cfscript>
