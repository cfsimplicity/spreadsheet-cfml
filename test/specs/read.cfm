<cfscript>
describe( "read", ()=>{

	it( "Can read an XLS file into a workbook object", ()=>{
		var path = getTestFilePath( "test.xls" );
		var workbook = s.read( src=path );
		expect( s.isBinaryFormat( workbook ) ).toBeTrue();
		workbook = s.newChainable().read( path ).getWorkbook();
		expect( s.isBinaryFormat( workbook ) ).toBeTrue();
	})

	it( "Can read an XLSX file into a workbook object", ()=>{
		var path = getTestFilePath( "test.xlsx" );
		var workbook = s.read( src=path );
		expect( s.isXmlFormat( workbook ) ).toBeTrue();
		workbook = s.newChainable().read( path ).getWorkbook();
		expect( s.isXmlFormat( workbook ) ).toBeTrue();
	})

	it( "Chainable method ends the chain and returns the import result if format is specified", ()=>{
		variables.spreadsheetTypes.Each( ( type )=>{
			var path = getTestFilePath( "test.#type#" );
			var data = s.newChainable().read( path, "query" );
			expect( IsQuery( data ) ).toBeTrue();
		})
	})

	it( "can read a spreadsheet into a query, array or array of structs", ()=>{
		var data = [ [ "Frumpo McNugget", "12345" ] ];
    spreadsheetTypes.Each( ( type )=>{
			var path = variables[ "temp" & type & "Path" ];
			s.newChainable( type )
        .addRows( data )
				.write( path, true );
			//array
			var expected = { columns: [], data: data };
			var actual = s.read( src=path, format="array" );
      expect( actual ).toBe( expected );
			//array of structs
			expected = [ [ column1: "Frumpo McNugget", column2:"12345" ] ];
			actual = s.read( src=path, format="arrayOfStructs" );
			expect( actual ).toBe( expected );
			//query
			expected = QueryNew( "column1,column2", "VarChar,VarChar", data );
			actual = s.read( src=path, format="query");
      expect( actual ).toBe( expected );
    })
  })

	it( "Returns no data if there are no *visible* sheets", ()=>{
		spreadsheetTypes.Each( ( type )=>{
			var path = variables[ "temp" & type & "Path" ];
			s.newChainable( type )
				.renameSheet( "hidden sheet", 1 )
				.setCellValue( "I'm in a hidden sheet", 1, 1 )
				.hideSheet( sheetNumber=1 )
			  .write( path, true );
			//array
			var expected = [ columns: [], data: [] ];
			var actual = s.read( src=path, format="array" );
			expect( actual ).toBe( expected );
			//query
			expected = QueryNew( "" );
			actual = s.read( src=path, format="query" );
			expect( actual ).toBe( expected );
		})
	})

	it( "Uses the first *visible* sheet if no sheet specified", ()=>{
		spreadsheetTypes.Each( ( type )=>{
			var path = variables[ "temp" & type & "Path" ];
			s.newChainable( type )
				.renameSheet( "hidden sheet", 1 )
				.setCellValue( "I'm in a hidden sheet", 1, 1 )
				.createSheet( "visible sheet" )
				.setActiveSheetNumber( 2 )
				.setCellValue( "I'm in a visible sheet", 1, 1 )
				.hideSheet( sheetNumber=1 )
				.write( path, true );
			//array
			var expected = [ columns: [], data: [ [ "I'm in a visible sheet" ] ] ];
			var actual = s.read( src=path, format="array" );
			expect( actual ).toBe( expected );
			//query
			expected = QueryNew( "column1", "VarChar", [ [ "I'm in a visible sheet" ] ] );
			actual = s.read( src=path, format="query" );
			expect( actual ).toBe( expected );
		})
	})

	describe( "read with headerRow", ()=>{

		it( "Uses the specified header row for column names", ()=>{
			var columns = [ "name", "number" ];
			var data = [ [ "Frumpo McNugget", "12345" ] ];
			spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRow( columns )
					.addRows( data )
					.write( path, true );
				//array
				var expected = { columns: columns, data: data };
				var actual = s.read( src=path, format="array", headerRow=1 );
				expect( actual ).toBe( expected );
				//arry of structs
				expected = [ [ name: "Frumpo McNugget", number: "12345" ] ];
				actual = s.read( src=path, format="arrayOfStructs", headerRow=1 );
				expect( actual ).toBe( expected );
				//query
				expected = QueryNew( columns.ToList(), "VarChar,VarChar", data );
				actual = s.read( src=path, format="query", headerRow=1 );
				expect( actual ).toBe( expected );
			})
		})

		it( "Generates default column names if the data has more columns than the specifed header row", ()=>{
			var headerRow = [ "firstColumn" ];
			var dataRow1 = [ "row 1 col 1 value" ];
			var dataRow2 = [ "row 2 col 1 value", "row 2 col 2 value" ];
			spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRow( headerRow )
					.addRow( dataRow1 )
					.addRow( dataRow2 )
					.write( path, true );
				//array
				var expected = { columns: [ "firstColumn", "column2" ], data: [ dataRow1, dataRow2 ] };
				var actual = s.read( src=path, format="array", headerRow=1 );
				expect( actual ).toBe( expected );
				//query
				expected = QueryNew( "firstColumn,column2", "VarChar,VarChar", [ dataRow1, dataRow2 ] );
				actual = s.read( src=path, format="query", headerRow=1 );
				expect( actual ).toBe( expected );
			})
		})

		it( "Includes the specified header row if includeHeader is true", ()=>{
			var columns = [ "name", "number" ];
			var data = [ "Frumpo McNugget", "12345" ];
			spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRow( columns )
					.addRow( data )
					.write( path, true );
				//array
				var expected = { columns: columns, data: [ columns, data ] };
				var actual = s.read( src=path, format="array", headerRow=1, includeHeaderRow=true );
				expect( actual ).toBe( expected );
				expected = QueryNew( columns.ToList(), "VarChar,VarChar", [ columns, data ] );
				actual = s.read( src=path, format="query", headerRow=1, includeHeaderRow=true );
				expect( actual ).toBe( expected );
				//query
			})
		})

		it( "Can handle null/empty cells", ()=>{
			var columns = [ "column1", "column2" ];
			var data = [ [ "", "a" ] ];
			spreadsheetTypes.Each( ( type )=>{
				var path = getTestFilePath( "nullCell." & type );
				//array
				var expected = { columns: columns , data: data };
				var actual = s.read( src=path, format="array", headerRow=1 );
				expect( actual ).toBe( expected );
				//query
				expected = QueryNew( columns.ToList(), "VarChar,VarChar", data );
				actual = s.read( src=path, format="query", headerRow=1 );
				expect( actual ).toBe( expected );
			})
		})

		it( "Includes trailing empty columns when using a header row", ()=>{
			var columns = [ "column1", "emptyColumn" ];
			var data = [ [ "column 1 value" ] ];
			spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRow( columns )
					.addRows( data )
					.write( path, true );
				//array
				var expected = { columns: columns, data: data };
				var actual = s.read( src=path, format="array", headerRow=1 );
				expect( actual ).toBe( expected );
				//query
				expected = QueryNew( columns.ToList(), "VarChar,VarChar", data );
				actual = s.read( src=path, format="query", headerRow=1 );
				expect( actual ).toBe( expected );
			})
		})

	})

	describe( "read with sheetName or sheetNumber", ()=>{

		it( "Reads from the specified sheet name or number", ()=>{
			var data = [ [ "Frumpo McNugget", "12345" ] ];
			spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.createSheet( "sheet2" )
					.setActiveSheet( "sheet2" )
					.addRows( data )
					.write( path, true );
				//array
				var expected = { columns: [], data: data };
				var actual = s.read( src=path, format="array", sheetName="sheet2" );
				expect( actual ).toBe( expected );
				actual = s.read( src=path, format="array", sheetNumber=2 );
				expect( actual ).toBe( expected );
				//query
				expected = QueryNew( "column1,column2", "VarChar,VarChar", data );
				var actual = s.read( src=path, format="query", sheetName="sheet2" );
				expect( actual ).toBe( expected );
				actual = s.read( src=path, format="query", sheetNumber=2 );
				expect( actual ).toBe( expected );
			})
		})

	})

	describe( "read with includeBlankRows", ()=>{

		it( "Excludes null and blank rows by default", ()=>{
			var data = [ [ "", "" ], [ "a", "b" ] ];
			spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRows( data )
					.write( path, true );
				//array
				var expected = { columns: [], data: [ [ "a", "b" ] ] };
				var actual = s.read( src=path, format="array" );
				expect( actual ).toBe( expected );
				//query
				expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ] ] );
				actual = s.read( src=path, format="query" );
				expect( actual ).toBe( expected );
			})
		})
	
		it( "Can include null and blank rows ", ()=>{
			var data = [ [ "", "" ], [ "a", "b" ] ];
			spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRows( data )
					.write( path, true );
				//array
				var expected = { columns: [], data: data };
				var actual = s.read( src=path, format="array", includeBlankRows=true );
				expect( actual ).toBe( expected );
				//query
				expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "" ], [ "a", "b" ] ] );
				var actual = s.read( src=path, format="query", includeBlankRows=true );
				expect( actual ).toBe( expected );
			})
		})

	})

	describe( "read with includeHiddenRows and includeHiddenColumns", ()=>{

		it( "Includes columns formatted as 'hidden' by default", ()=>{
			var columns = [ "col1", "col2" ];
			var data = [ [ "a1", "b1" ] ];
			spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRow( columns )
					.addRows( data )
					.hideColumn( 1 )
					.write( path, true );
				//array
				var actual = s.read( src=path, format="array", headerRow=1 );
				var expected = { columns: columns, data: data };
				expect( actual ).toBe( expected );
				//query
				actual = s.read( src=path, format="query", headerRow=1 );
				expected = QueryNew( "col1,col2", "VarChar,VarChar", [ [ "a1", "b1" ] ] );
				expect( actual ).toBe( expected );
			})
		})
	
		it( "Can exclude columns formatted as 'hidden'", ()=>{
			var columns = [ "col1", "col2" ];
			var data = [ [ "a1", "b1" ] ];
			spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRow( columns )
					.addRows( data )
					.hideColumn( 1 )
					.write( path, true );
				//array
				var actual = s.read( src=path, format="array", headerRow=1, includeHiddenColumns=false );
				var expected = { columns: [ "col2" ], data: [ [ "b1" ] ] };
				expect( actual ).toBe( expected );
				//query
				actual = s.read( src=path, format="query", headerRow=1, includeHiddenColumns=false );
				expected = QueryNew( "col2", "VarChar", [ [ "b1" ] ]  );
				expect( actual ).toBe( expected );
			})
		})
	
		it( "Includes rows formatted as 'hidden' by default", ()=>{
			var data = [ [ "Apple" ], [ "Banana" ], [ "Carrot" ], [ "Doughnut" ] ];
			spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRows( data )
					.hideRow( 1 )
					.write( path, true );
				//array
				var actual = s.read( src=path, format="array" );
				var expected = { columns: [], data: data };
				expect( actual ).toBe( expected );
				//query
				actual = s.read( src=path, format="query" );
				expected = QueryNew( "column1", "VarChar", data );
				expect( actual ).toBe( expected );
			})
		})
	
		it( "Can exclude rows formatted as 'hidden'", ()=>{
			var data = [ [ "Apple" ], [ "Banana" ], [ "Carrot" ], [ "Doughnut" ] ];
			spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRows( data )
					.hideRow( 1 )
					.hideRow( 3 )
					.write( path, true );
				//array
				var actual = s.read( src=path, format="array", includeHiddenRows=false );
				var expected = { columns: [], data: [ [ "Banana" ], [ "Doughnut" ] ] };
				expect( actual ).toBe( expected );
				//query
				actual = s.read( src=path, format="query", includeHiddenRows=false );
				expected = QueryNew( "column1", "VarChar", [ [ "Banana" ], [ "Doughnut" ] ] );
				expect( actual ).toBe( expected );
			})
		})

		it( "Returns an empty data set if excluding hidden columns and ALL columns are hidden", ()=>{
			spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addColumn( "a1" )
					.addColumn( "b1" )
					.hideColumn( 1 )
					.hideColumn( 2 )
					.write( path, true );
				//array
				var actual = s.read( src=path, format="array", includeHiddenColumns=false );
				var expected = { columns: [], data: [] };
				expect( actual ).toBe( expected );
				//query
				actual = s.read( src=path, format="query", includeHiddenColumns=false );
				expected = QueryNew( "" );
				expect( actual ).toBe( expected );
			})
			
		})

	})

	describe( "read with returnVisibleValues", ()=>{

		it( "Reads values of different types correctly, by default returning the raw values", ()=>{
			var numericValue = 2;
			var dateValue = CreateDate( 2015, 04, 12 );
			var rawDecimalValue = 0.000011;
			var leadingZeroValue = "01";
			var columns = [ "numeric", "zero", "decimal", "boolean", "date", "leadingZero" ];
			var data = [ [ numericValue, 0, rawDecimalValue, true, dateValue, leadingZeroValue ] ];
			variables.spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRows( data )
					.formatCell( { dataformat: "0.00000" }, 1, 3 )
					.write( path, true );
				//array
				var expected = { columns: columns, data: data };
				var actual = s.read( src=path, format="array", columnNames=columns );
				expect( actual ).toBe( expected );
				//query
				expected = QueryNew( columns.ToList(), "Integer,Integer,Decimal,Bit,Date,VarChar", [ [ numericValue, 0, rawDecimalValue, true, dateValue, leadingZeroValue ] ] );
				actual = s.read( src=path, format="query", columnNames=columns );
				expect( actual ).toBe( expected );
			})
		})
	
		it( "Can return the visible/formatted value rather than raw value", ()=>{
			var rawDecimalValue = 0.000011;
			var visibleDecimalValue = 0.00001;
			var data = [ [ rawDecimalValue ] ];
			variables.spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRows( data )
					.formatCell( { dataformat: "0.00000" }, 1, 1 )
					.write( path, true );
				//array
				var expected = { columns: [], data: [ [ visibleDecimalValue ] ] };
				var actual = s.read( src=path, format="array", returnVisibleValues=true );
				expect( actual ).toBe( expected );
				var decimalHasBeenOutputInScientificNotation = ( Trim( actual.data[ 1 ][ 1 ] ).FindNoCase( "E" ) > 0 );
				expect( decimalHasBeenOutputInScientificNotation ).toBeFalse();
				//query
				actual = s.read( src=path, format="query", returnVisibleValues=true );
				expect( actual.column1 ).toBe( visibleDecimalValue );
				decimalHasBeenOutputInScientificNotation = ( Trim( actual.column1 ).FindNoCase( "E" ) > 0 );
				expect( decimalHasBeenOutputInScientificNotation ).toBeFalse();
			})
		})

	})

	describe( "read with fillMergedCellsWithVisibleValue", ()=>{

		it( "Can fill each of the empty cells in merged regions with the visible merged cell value without conflicting with includeBlankRows=true", ()=>{
			var data = [ [ "a", "b" ], [ "c", "d" ], [ "", "" ] ];
			variables.spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRows( data )
					.mergeCells( 1, 2, 1, 2, true )//force empty merged cells
					.write( path, true );
				//array
				var expected = { columns: [], data: [ [ "a", "a" ], [ "a", "a" ] ] };
				var actual = s.read( src=path, format="array", fillMergedCellsWithVisibleValue=true );
				expect( actual ).toBe( expected );
				//test retention of blank row not part of merge region
				expected = { columns: [], data: [ [ "a", "a" ], [ "a", "a" ], [ "", "" ] ] };
				actual = s.read( src=path, format="array", fillMergedCellsWithVisibleValue=true, includeBlankRows=true );
				expect( actual ).toBe( expected );
				//query
				expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "a" ], [ "a", "a" ] ] );
				actual = s.read( src=path, format="query", fillMergedCellsWithVisibleValue=true );
				expect( actual ).toBe( expected );
				//test retention of blank row not part of merge region
				expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "a" ], [ "a", "a" ], [ "", "" ] ] );
				actual = s.read( src=path, format="query", fillMergedCellsWithVisibleValue=true, includeBlankRows=true );
				expect( actual ).toBe( expected );
			})
		})
		
	})

	describe( "read with password encryption", ()=>{
		
		it( "Can read an encrypted file", ()=>{
			spreadsheetTypes.Each( ( type )=>{
				var path = getTestFilePath( "passworded." & type );
				//array
				var expected = { columns: [], data: [ [ "secret" ] ] };
				var actual = s.read( src=path, format="array", password="pass" );
				expect( actual ).toBe( expected );
				//query
				expected = QueryNew( "column1", "VarChar", [ [ "secret" ] ] );
				actual = s.read( src=path, format="query", password="pass" );
				expect( actual ).toBe( expected );
			})
		})

	})

	describe( "read specific rows and/or columns", ()=>{

		it( "Can read specific rows only", ()=>{
			var data = [ [ "row1" ], [ "row2" ] ];
			variables.spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRows( data )
					.write( path, true );
				//array
				var actual = s.read( src=path, format="array", rows="2" );
				var expected = { columns: [], data: [ [ "row2" ] ] };
				expect( actual ).toBe( expected );
				//query
				actual = s.read( src=path, format="query", rows="2" );
				expected = QueryNew( "column1", "VarChar", [ [ "row2" ] ] );
				expect( actual ).toBe( expected );
			})
		})

		it( "Can read specific columns only", ()=>{
			var data = [ [ "a", "b" ] ];
			variables.spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRows( data )
					.write( path, true );
				//array
				var actual = s.read( src=path, format="array", columns="2" );
				var expected = { columns: [], data: [ [ "b" ] ] };
				expect( actual ).toBe( expected );
				//query
				actual = s.read( src=path, format="query", columns="2" );
				expected = QueryNew( "column1", "VarChar", [ [ "b" ] ] );
				expect( actual ).toBe( expected );
			})
		})

		it( "Can read specific rows and columns only", ()=>{
			var data = QuerySim( "A1,B1,C1,D1,E1
				A2|B2|C2|D2|E2
				A3|B3|C3|D3|E3
				A4|B4|C4|D4|E4
				A5|B5|C5|D5|E5");
			variables.spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.fromQuery( data, true )
					.write( path, true );
				var actual = s.read( src=path, format="query", columns="2,4-5", rows="2,4-5", headerRow=1 );
				var expected = QuerySim( "B1,D1,E1
					B2|D2|E2
					B4|D4|E4
					B5|D5|E5");
				expect( actual ).toBe( expected );
			})
		})

		it( "Can read data starting at specific rows and/or columns", ()=>{
			var data = QuerySim( "A1,B1,C1,D1,E1,F1
				A2|B2|C2|D2|E2|F2
				A3|B3|C3|D3|E3|F3
				A4|B4|C4|D4|E4|F4
				A5|B5|C5|D5|E5|F5
				A6|B6|C6|D6|E6|F6");
			variables.spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.fromQuery( data, true )
					.write( path, true );
				var actual = s.read( src=path, format="query", columns="2,4-", rows="2,4-", headerRow=1 );
				var expected = QuerySim( "B1,D1,E1,F1
					B2|D2|E2|F2
					B4|D4|E4|F4
					B5|D5|E5|F5
					B6|D6|E6|F6");
				expect( actual ).toBe( expected );
			})
		})

	})

	describe( "read with columnNames", ()=>{

		it( "Returns column names but no data if column names are specified but spreadsheet is empty", ()=>{
			variables.spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type ).write( path, true );
				//array
				var expected = { columns: [ "One", "Two" ], data: [] };
				var actual = s.read( src=path, format="array", columnNames="One,Two" );
				expect( actual ).toBe( expected );
				//query
				expected = QueryNew( "One,Two","Varchar,Varchar", [] );
				actual = s.read( src=path, format="query", columnNames="One,Two" );
				expect( actual ).toBe( expected );
			})
		})

		it( "Allows column names to be specified as a list or array when reading a sheet", ()=>{
			var columns = [ "Name", "Number" ];
			var data = [ [ "Frumpo McNugget", "12345" ] ];
			variables.spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRows( data )
					.write( path, true );
				//array
				var expected = { columns: columns, data: data };
				var actual = s.read( src=path, format="array", columnNames="Name,Number" ); //list
				expect( actual ).toBe( expected );
				actual = s.read( src=path, format="array", columnNames=columns ); //array
				expect( actual ).toBe( expected );
				//query
				expected= QueryNew( "Name,Number", "VarChar,VarChar", data );
				actual = s.read( src=path, format="query", columnNames="Name,Number" ); // list
				expect( actual ).toBe( expected );
				actual = s.read( src=path, format="query", columnNames=columns );// array
				expect( actual ).toBe( expected );
			})
		})

		it( "ColumnNames overrides headerRow: none of the header row values will be used", ()=>{
			var columns = [ "Name", "Number" ];
			var data = [ [ "Frumpo McNugget", "12345" ] ];
			variables.spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRow( [ "Col1", "Col2" ] ) //this will be ignored
					.addRows( data )
					.write( path, true );
				//array
				var expected = { columns: columns, data: data };
				var actual = s.read( src=path, format="array", columnNames=columns, headerRow=1 ); //array
				expect( actual ).toBe( expected );
				//query
				expected= QueryNew( "Name,Number", "VarChar,VarChar", data );
				actual = s.read( src=path, format="query", columnNames=columns, headerRow=1 );// array
				expect( actual ).toBe( expected );
			})
		})

		it( "can handle column names containing commas or spaces", ()=>{
			var path = getTestFilePath( "commaAndSpaceInColumnHeaders.xls" );
			var actual = s.read( src=path, format="query", headerRow=1 );
			var columnNames = [ "first name", "surname,comma" ];// these are the file column headers
			expect( actual.getColumnNames()[ 1 ] ).toBe( columnNames[ 1 ] );
			expect( actual.getColumnNames()[ 2 ] ).toBe( columnNames[ 2 ] );
		})

		it( "Accepts 'queryColumnNames' as an alias of 'columnNames'", ()=>{
			var data = [ [ "Frumpo McNugget", "12345" ] ];
			variables.spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRows( data )
					.write( path, true );
				var expected = QueryNew( "One,Two", "VarChar,VarChar", data );
				var actual = s.read( src=path, format="query", queryColumnNames="One,Two" );
				expect( actual ).toBe( expected );
			})
		})

		it( "Generates default column names if the data has more columns than the specifed column names", ()=>{
			var columnNames = [ "firstColumn" ];
			var dataRow1 = [ "row 1 col 1 value" ];
			var dataRow2 = [ "row 2 col 1 value", "row 2 col 2 value" ];
			variables.spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRow( dataRow1 )
					.addRow( dataRow2 )
					.write( path, true );
				//array
				var expected = { columns: [ "firstColumn", "column2" ], data: [ dataRow1, dataRow2 ] };
				var actual = s.read( src=path, format="array", columnNames=columnNames );
				expect( actual ).toBe( expected );
				//query
				expected = QueryNew( "firstColumn,column2", "VarChar,VarChar", [ dataRow1, dataRow2 ] );
				var actual = s.read( src=path, format="query", columnNames=columnNames );
				expect( actual ).toBe( expected );
			})
		})

	})

})
</cfscript>