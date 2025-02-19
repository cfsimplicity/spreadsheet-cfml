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

	it( "Can read a traditional XLS file into a query", ()=>{
		var path = getTestFilePath( "test.xls" );
		var expected = querySim(
			"column1,column2
			a|b
			1|#CreateDateTime( 2015, 4, 1, 0, 0, 0 )#
			#CreateDateTime( 2015, 4, 1, 1, 1, 1 )#|2");
		var actual = s.read( src=path,format="query" );
		expect( actual ).toBe( expected );

	})

	it( "Can read an OOXML file into a query", ()=>{
		var path = getTestFilePath( "test.xlsx" );
		var expected = querySim(
			"column1,column2
			a|e
			b|f
			c|g
			I am|ooxml");
		var actual = s.read( src=path, format="query" );
	})

	it( "Uses the first *visible* sheet if format=query and no sheet specified", ()=>{
		spreadsheetTypes.Each( ( type )=>{
			var path = variables[ "temp" & type & "Path" ];
			var wb = s.newChainable( type )
				.renameSheet( "hidden sheet", 1 )
				.setCellValue( "I'm in a hidden sheet", 1, 1 )
				.createSheet( "visible sheet" )
				.setActiveSheetNumber( 2 )
				.setCellValue( "I'm in a visible sheet", 1, 1 )
				.getWorkbook();
			s.getSheetHelper().setVisibility( wb, 1, "VERY_HIDDEN" );
			s.write( wb, path, true );
			var expected = QueryNew( "column1", "VarChar", [ [ "I'm in a visible sheet" ] ] );
			var actual = s.read( src=path, format="query" );
			expect( actual ).toBe( expected );
		})
	})

	it( "Returns a blank query if format=query and there are no visible sheets", ()=>{
		spreadsheetTypes.Each( ( type )=>{
			var path = variables[ "temp" & type & "Path" ];
			var wb = s.newChainable( type )
				.renameSheet( "hidden sheet", 1 )
				.setCellValue( "I'm in a hidden sheet", 1, 1 )
				.getWorkbook();
			s.getSheetHelper().setVisibility( wb, 1, "VERY_HIDDEN" );
			s.write( wb, path, true );
			var expected = QueryNew( "" );
			var actual = s.read( src=path, format="query" );
			expect( actual ).toBe( expected );
		})
	})

	it( "Reads from the specified sheet name", ()=>{
		var path = getTestFilePath( "test.xls" );// has 2 sheets
		var expected = querySim(
			"column1,column2
			x|y");
		var actual = s.read( src=path, format="query", sheetName="sheet2" );
		expect( actual ).toBe( expected );
	})

	it( "Reads from the specified sheet number", ()=>{
		var path = getTestFilePath( "test.xls" );// has 2 sheets
		var expected = querySim(
			"column1,column2
			x|y");
		var actual = s.read( src=path, format="query", sheetNumber=2 );
		expect( actual ).toBe( expected );
	})

	it( "Uses the specified header row for column names", ()=>{
		var path = getTestFilePath( "test.xls" );
		var expected = querySim(
			"a,b
			1|#CreateDateTime( 2015, 4, 1, 0, 0, 0 )#
			#CreateDateTime( 2015, 4, 1, 1, 1, 1 )#|2");
		var actual = s.read( src=path, format="query", headerRow=1 );
		expect( actual ).toBe( expected );
	})

	it( "Generates default column names if the data has more columns than the specifed header row", ()=>{
		var headerRow = [ "firstColumn" ];
		var dataRow1 = [ "row 1 col 1 value" ];
		var dataRow2 = [ "row 2 col 1 value", "row 2 col 2 value" ];
		var expected = querySim(
			"firstColumn,column2
			row 1 col 1 value|
			row 2 col 1 value|row 2 col 2 value"
		);
		s.newChainable( "xls" )
		 .addRow( headerRow )
		 .addRow( dataRow1 )
		 .addRow( dataRow2 )
		 .write( tempXlsPath, true );
		var actual = s.read( src=tempXlsPath, format="query", headerRow=1 );
		expect( actual ).toBe( expected );
	})

	it( "Includes the specified header row in query if includeHeader is true", ()=>{
		var path = getTestFilePath( "test.xls" );
		var expected = querySim(
			"a,b
			a|b
			1|#CreateDateTime( 2015, 4, 1, 0, 0, 0 )#
			#CreateDateTime( 2015, 4, 1, 1, 1, 1 )#|2");
		var actual = s.read( src=path, format="query", headerRow=1, includeHeaderRow=true );
		expect( actual ).toBe( expected );
	})

	it( "Excludes null and blank rows in query by default", ()=>{
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "" ], [ "a", "b" ] ] );
		var workbook = s.new();
		s.addRows( workbook, data )
			.write( workbook, tempXlsPath, true );
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ] ] );
		var actual = s.read( src=tempXlsPath, format="query" );
		expect( actual ).toBe( expected );
	})

	it( "Includes null and blank rows in query if includeBlankRows is true", ()=>{
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "" ], [ "a", "b" ] ] );
		var workbook = s.new();
		s.addRows( workbook, data )
			.write( workbook, tempXlsPath, true );
		var expected = data;
		var actual = s.read( src=tempXlsPath, format="query", includeBlankRows=true );
		expect( actual ).toBe( expected );
	})

	it( "Can handle null/empty cells", ()=>{
		var path = getTestFilePath( "nullCell.xls" );
		var actual = s.read( src=path, format="query", headerRow=1 );
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "a" ] ] );
		expect( actual ).toBe( expected );
	})

	it( "Includes trailing empty columns when using a header row", ()=>{
		var paths = [ tempXlsPath, tempXlsxPath ];
		var expected = QuerySim( "col1,col2,emptyCol
			value|value|");
		paths.Each( ( path )=>{
			var type = ( path == tempXlsPath )? "xls": "xlsx";
			var workbook = s.newChainable( type )
				.addRow( "col1,col2,emptyCol" )
				.addRow( "value,value" )
				.write( path, true );
			var actual = s.read( src=path, format="query", headerRow=1 );
			expect( actual ).toBe( expected );
		})
	})

	it( "Reads values of different types correctly, by default returning the raw values", ()=>{
		var numericValue = 2;
		var dateValue = CreateDate( 2015, 04, 12 );
		var rawDecimalValue = 0.000011;
		var leadingZeroValue = "01";
		var columnNames = [ "numeric", "zero", "decimal", "boolean", "date", "leadingZero" ];
		var data = QueryNew( columnNames.ToList(), "Integer,Integer,Decimal,Bit,Date,VarChar", [ [ numericValue, 0, rawDecimalValue, true, dateValue, leadingZeroValue ] ] );
		variables.spreadsheetTypes.Each( ( type )=>{
			var path = ( type == "xls" )? tempXlsPath: tempXlsxPath;
			s.newChainable( type )
				.addRows( data )
				.formatCell( { dataformat: "0.00000" }, 1, 3 )
				.write( path, true );
			var actual = s.read( src=path, format="query", columnNames=columnNames );
			expect( actual ).toBe( data );
		})
	})

	
	it( "Can return the visible/formatted values rather than raw values", ()=>{
		var numericValue = 2;
		var dateValue = CreateDate( 2015, 04, 12 );
		var rawDecimalValue = 0.000011;
		var visibleDecimalValue = 0.00001;
		var leadingZeroValue = "01";
		var columnNames = [ "numeric", "zero", "decimal", "boolean", "date", "leadingZero" ];
		var data = QueryNew( columnNames.ToList(), "Integer,Integer,Decimal,Bit,Date,VarChar", [ [ numericValue, 0, rawDecimalValue, true, dateValue, leadingZeroValue ] ] );
		variables.spreadsheetTypes.Each( ( type )=>{
			var path = ( type == "xls" )? tempXlsPath: tempXlsxPath;
			s.newChainable( type )
				.addRows( data )
				.formatCell( { dataformat: "0.00000" }, 1, 3 )
				.write( path, true );
			var actual = s.read( src=path, format="query", columnNames=columnNames, returnVisibleValues=true );
			expect( actual.numeric ).toBe( numericValue );
			expect( actual.zero ).toBe( 0 );
			expect( actual.decimal ).toBe( visibleDecimalValue );
			expect( actual.boolean ).toBeTrue();
			expect( actual.date ).toBe( dateValue );
			expect( actual.leadingZero ).toBe( leadingZeroValue );
			var decimalHasBeenOutputInScientificNotation = ( Trim( actual.decimal ).FindNoCase( "E" ) > 0 );
			expect( decimalHasBeenOutputInScientificNotation ).toBeFalse();
		})
	})
	

	it( "Can fill each of the empty cells in merged regions with the visible merged cell value without conflicting with includeBlankRows=true", ()=>{
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ], [ "", "" ] ] );
		var workbook = s.workbookFromQuery( data, false );
		s.mergeCells( workbook, 1, 2, 1, 2, true )//force empty merged cells
			.write( workbook, tempXlsPath, true );
		var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "a" ], [ "a", "a" ] ] );
		var actual = s.read( src=tempXlsPath, format="query", fillMergedCellsWithVisibleValue=true );
		expect( actual ).toBe( expected );
		//test retention of blank row not part of merge region
		expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "a" ], [ "a", "a" ], [ "", "" ] ] );
		actual = s.read( src=tempXlsPath, format="query", fillMergedCellsWithVisibleValue=true, includeBlankRows=true );
		expect( actual ).toBe( expected );
	})

	it( "Can read specified rows only into a query", ()=>{
		var data = QuerySim( "A
			A1
			A2
			A3
			A4
			A5");
		var workbook = s.workbookFromQuery( data, false );
		s.write( workbook, tempXlsPath, true );
		var actual = s.read( src=tempXlsPath, format="query", rows="2,4-5" );
		var expected = QuerySim( "column1
			A2
			A4
			A5");
		expect( actual ).toBe( expected );
		//with header row included in row 1
		data = QuerySim( "A1
			A2
			A3
			A4
			A5
			A6");
		var workbook = s.workbookFromQuery( data, true );
		s.write( workbook, tempXlsPath, true );
		var actual = s.read( src=tempXlsPath, format="query", rows="2,4-5", headerRow=1 );
		expected = QuerySim( "A1
			A2
			A4
			A5");
		expect( actual ).toBe( expected );
	})

	it( "Can read specified column numbers only into a query", ()=>{
		var data = QuerySim( "A,B,C,D,E
			A1|B1|C1|D1|E1");
		//With no header row, so no column names specified
		var workbook = s.workbookFromQuery( data, false );
		s.write( workbook, tempXlsPath, true );
		var actual = s.read( src=tempXlsPath, format="query", columns="2,4-5" );
		var expected = QuerySim( "column1,column2,column3
			B1|D1|E1");
		expect( actual ).toBe( expected );
		//With column names specified from the header row
		var workbook = s.workbookFromQuery( data, true );
		s.write( workbook ,tempXlsPath, true );
		actual = s.read( src=tempXlsPath, format="query", columns="2,4-5", headerRow=1 );
		expected = QuerySim( "B,D,E
			B1|D1|E1");
		expect( actual ).toBe( expected );
		//Include the header row with specified column names
		var workbook = s.workbookFromQuery( data, true );
		s.write( workbook, tempXlsPath, true );
		actual = s.read( src=tempXlsPath, format="query", columns="2,4-5", headerRow=1, includeHeaderRow=true );
		expected = QuerySim( "B,D,E
			B|D|E
			B1|D1|E1");
		expect( actual ).toBe( expected );
	})

	it( "Can read specific rows and columns only into a query", ()=>{
		var data = QuerySim( "A1,B1,C1,D1,E1
			A2|B2|C2|D2|E2
			A3|B3|C3|D3|E3
			A4|B4|C4|D4|E4
			A5|B5|C5|D5|E5");
		//First row is header
		var workbook = s.workbookFromQuery( data, true );
		s.write( workbook, tempXlsPath, true );
		var actual = s.read( src=tempXlsPath, format="query", columns="2,4-5", rows="2,4-5", headerRow=1 );
		var expected = QuerySim( "B1,D1,E1
			B2|D2|E2
			B4|D4|E4
			B5|D5|E5");
		expect( actual ).toBe( expected );
	})

	it( "Can read data starting at specific rows and/or columns into a query", ()=>{
		var data = QuerySim( "A1,B1,C1,D1,E1,F1
			A2|B2|C2|D2|E2|F2
			A3|B3|C3|D3|E3|F3
			A4|B4|C4|D4|E4|F4
			A5|B5|C5|D5|E5|F5
			A6|B6|C6|D6|E6|F6");
		var workbook = s.workbookFromQuery( data=data, addHeaderRow=true );
		s.write( workbook, tempXlsPath, true );
		workbook = s.workbookFromQuery( data=data, addHeaderRow=true, xmlFormat=true );
		s.write( workbook, tempXlsxPath, true );
		var paths = [ tempXlsPath, tempXlsxPath ];
		paths.Each( ( path )=>{
			var actual = s.read( src=path, format="query", columns="2,4-", rows="2,4-", headerRow=1 );
			var expected = QuerySim( "B1,D1,E1,F1
				B2|D2|E2|F2
				B4|D4|E4|F4
				B5|D5|E5|F5
				B6|D6|E6|F6");
			expect( actual ).toBe( expected );
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

	it( "Returns an empty query if the spreadsheet is empty even if headerRow is specified", ()=>{
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( ( wb )=>{
			var path = s.isXmlFormat( wb )? tempXlsxPath: tempXlsPath;
			s.write( wb, path, true );
			var actual = s.read( src=path, format="query", headerRow=1 );
			var expected = QueryNew( "" );
			expect( actual ).toBe( expected );
		})
	})

	it( "Returns an empty query if excluding hidden columns and ALL columns are hidden", ()=>{
		var workbook = s.new();
		s.addColumn( workbook, "a1" )
			.addColumn( workbook, "b1" )
			.hideColumn( workbook, 1 )
			.hideColumn( workbook, 2 )
			.write( workbook, tempXlsPath, true );
		var actual = s.read( src=tempXlsPath, format="query", includeHiddenColumns=false );
		var expected = QueryNew( "" );
		expect( actual ).toBe( expected );
	})

	it( "Returns a query with column names but no rows if column names are specified but spreadsheet is empty", ()=>{
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( ( wb )=>{
			var path = s.isXmlFormat( wb )? tempXlsxPath: tempXlsPath;
			s.write( wb, path, true );
			var actual = s.read( src=path, format="query", columnNames="One,Two" );
			var expected = QueryNew( "One,Two","Varchar,Varchar", [] );
			expect( actual ).toBe( expected );
		})
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

	describe( "query column name setting", ()=>{

		it( "Allows column names to be specified as a list when reading a sheet into a query", ()=>{
			var path = getTestFilePath( "test.xls" );
			actual = s.read( src=path, format="query", columnNames="One,Two" );
			expected = QuerySim( "One,Two
				a|b
				1|#CreateDateTime( 2015, 4, 1, 0, 0, 0 )#
				#CreateDateTime( 2015, 4, 1, 1, 1, 1 )#|2");
			expect( actual ).toBe( expected );
		})

		it( "Allows column names to be specified as an array when reading a sheet into a query", ()=>{
			var path = getTestFilePath( "test.xls" );
			actual = s.read( src=path, format="query", columnNames=[ "One", "Two" ] );
			expected = QuerySim( "One,Two
				a|b
				1|#CreateDateTime( 2015, 4, 1, 0, 0, 0 )#
				#CreateDateTime( 2015, 4, 1, 1, 1, 1 )#|2");
			expect( actual ).toBe( expected );
		})

		it( "ColumnNames list overrides headerRow: none of the header row values will be used", ()=>{
			var path = getTestFilePath( "test.xls" );
			var actual = s.read( src=path, format="query", columnNames="One,Two", headerRow=1 );
			var expected = QuerySim( "One,Two
				1|#CreateDateTime( 2015, 4, 1, 0, 0, 0 )#
				#CreateDateTime( 2015, 4, 1, 1, 1, 1 )#|2");
			expect( actual ).toBe( expected );
		})

		it( "can handle column names containing commas or spaces", ()=>{
			var path = getTestFilePath( "commaAndSpaceInColumnHeaders.xls" );
			var actual = s.read( src=path, format="query", headerRow=1 );
			var columnNames = [ "first name", "surname,comma" ];// these are the file column headers
			expect( actual.getColumnNames()[ 1 ] ).toBe( columnNames[ 1 ] );
			expect( actual.getColumnNames()[ 2 ] ).toBe( columnNames[ 2 ] );
		})

		it( "Accepts 'queryColumnNames' as an alias of 'columnNames'", ()=>{
			var path = getTestFilePath( "test.xls" );
			actual = s.read( src=path, format="query", queryColumnNames="One,Two" );
			expected = QuerySim( "One,Two
				a|b
				1|#CreateDateTime( 2015, 4, 1, 0, 0, 0 )#
				#CreateDateTime( 2015, 4, 1, 1, 1, 1 )#|2");
			expect( actual ).toBe( expected );
		})

		it( "Allows header names to be made safe for query column names", ()=>{
			var data = [ [ "id","id","A  B","x/?y","(a)"," A","##1","1a" ], [ 1,2,3,4,5,6,7,8 ] ];
			var wb = s.newXlsx();
			s.addRows( wb, data )
				.write( wb, tempXlsxPath, true );
			var q = s.read( src=tempXlsxPath, format="query", headerRow=1, makeColumnNamesSafe=true );
			var expected = [ "id", "id2", "A_B", "x_y", "_a_", "A", "Number1", "_a" ];
			cfloop( from=1, to=expected.Len(), index="i" ){
				expect( q.getColumnNames()[ i ] ).toBe( expected[ i ] );
			}
			var wb = s.newXls();
			s.addRows( wb, data )
				.write( wb, tempXlsPath, true );
			var q = s.read( src=tempXlsPath, format="query", headerRow=1, makeColumnNamesSafe=true );
			var expected = [ "id", "id2", "A_B", "x_y", "_a_", "A", "Number1", "_a" ];
			cfloop( from=1, to=expected.Len(), index="i" ){
				expect( q.getColumnNames()[ i ] ).toBe( expected[ i ] );
			}
		})

		it( "Generates default column names if the data has more columns than the specifed column names", ()=>{
			var columnNames = [ "firstColumn" ];
			var dataRow1 = [ "row 1 col 1 value" ];
			var dataRow2 = [ "row 2 col 1 value", "row 2 col 2 value" ];
			var expected = querySim(
				"firstColumn,column2
				row 1 col 1 value|
				row 2 col 1 value|row 2 col 2 value"
			);
			s.newChainable( "xls" )
			 .addRow( dataRow1 )
			 .addRow( dataRow2 )
			 .write( tempXlsPath, true );
			var actual = s.read( src=tempXlsPath, format="query", columnNames=columnNames );
			expect( actual ).toBe( expected );
		})

	})

	describe( "query column type setting", ()=>{

		it( "allows the query column types to be manually set using list", ()=>{
			var workbook = s.new();
			s.addRow( workbook, [ 1, 1.1, "string", _CreateTime( 1, 0, 0 ) ] )
				.write( workbook, tempXlsPath, true );
			var q = s.read( src=tempXlsPath, format="query", queryColumnTypes="Integer,Double,VarChar,Time" );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "INTEGER" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIME" );
		})

		it( "allows the query column types to be manually set where the column order isn't known, but the header row values are", ()=>{
			var workbook = s.new();
			s.addRows( workbook, [ [ "integer", "double", "string column", "time" ], [ 1, 1.1, "text", _CreateTime( 1, 0, 0 ) ] ] )
				.write( workbook, tempXlsPath, true );
			var columnTypes = { "string column": "VARCHAR", "integer": "INTEGER", "time": "TIME", "double": "DOUBLE" };//not in order
			var q = s.read( src=tempXlsPath, format="query", queryColumnTypes=columnTypes, headerRow=1 );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "INTEGER" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIME" );
		})

		it( "allows the query column types to be manually set where the column order isn't known, but the column names are", ()=>{
			var workbook = s.new();
			s.addRows( workbook, [ [ 1, 1.1, "text", _CreateTime( 1, 0, 0 ) ] ] )
				.write( workbook, tempXlsPath, true );
			var columnNames = "integer,double,string column,time";
			var columnTypes = { "string": "VARCHAR", "integer": "INTEGER", "time": "TIME", "double": "DOUBLE" };//not in order
			var q = s.read( src=tempXlsPath, format="query", queryColumnTypes=columnTypes, columnNames=columnNames );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "INTEGER" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIME" );
		})

		it( "allows the query column types to be automatically set", ()=>{
			var workbook = s.new();
			s.addRow( workbook, [ 1, 1.1, "string", Now() ] )
				.write( workbook, tempXlsPath, true );
			var q = s.read( src=tempXlsPath, format="query", queryColumnTypes="auto" );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIMESTAMP" );
		})

		it( "automatic detecting of query column types ignores blank cells", ()=>{
			var workbook = s.new();
			var data = [
				[ "", "", "", "" ],
				[ "", 2, "test", Now() ],
				[ 1, 1.1, "string", Now() ],
				[ 1, "", "", "" ]
			];
			s.addRows( workbook, data )
				.write( workbook, tempXlsPath, true );
			var q = s.read( src=tempXlsPath, format="query", queryColumnTypes="auto" );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIMESTAMP" );
		})

		it( "allows a default type to be set for all query columns", ()=>{
			var workbook = s.new();
			s.addRow( workbook, [ 1, 1.1, "string", Now() ] )
				.write( workbook, tempXlsPath, true );
			var q = s.read( src=tempXlsPath, format="query", queryColumnTypes="VARCHAR" );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 2 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "VARCHAR" );
		})

	})

	describe(
		title="Lucee only timezone tests",
		body=function(){

			it( "Doesn't offset a date value even if the Lucee timezone doesn't match the system", ()=>{
				variables.currentTZ = GetTimeZone();
				variables.tempTZ = "US/Eastern";
				spreadsheetTypes.Each( ( type )=>{
					SetTimeZone( tempTZ );
					var path = variables[ "temp" & type & "Path" ];
					local.s = newSpreadsheetInstance();//timezone mismatch detection cached is per instance
					local.s.newChainable( type ).setCellValue( "2022-01-01", 1, 1, "date" ).write( path, true );
					var actual = local.s.read( path, "query" ).column1;
					var expected = CreateDate( 2022, 01, 01 );
					expect( actual ).toBe( expected );
					SetTimeZone( currentTZ );
				})

			})

		},
		skip=( !s.getIsLucee() || ( s.getDateHelper().getPoiTimeZone() != "Europe/London" ) )// only valid if system timezone is ahead of temporary test timezone
	);

	describe( "read throws an exception if", ()=>{

		it( "queryColumnTypes is specified as a 'columnName/type' struct, but headerRow and columnNames arguments are missing", ()=>{
			expect( ()=>{
				var columnTypes = { col1: "Integer" };
				s.read( src=getTestFilePath( "test.xlsx" ), format="query", queryColumnTypes=columnTypes );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidQueryColumnTypesArgument" );
		})

		it( "a formula can't be evaluated", ()=>{
			expect( ()=>{
				var workbook = s.new();
				s.addColumn( workbook, "1,1" );
				var theFormula="SUS(A1:A2)";//invalid formula
				s.setCellFormula( workbook, theFormula, 3, 1 )
					.write( workbook=workbook, filepath=tempXlsPath, overwrite=true )
					.read( src=tempXlsPath, format="query" );
			}).toThrow( type="cfsimplicity.spreadsheet.failedFormula" );
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
