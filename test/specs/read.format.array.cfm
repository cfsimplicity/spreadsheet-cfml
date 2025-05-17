<cfscript>
describe( "read: format=array", ()=>{

  it( "can read a spreadsheet into an array", ()=>{
    var expected = { columns: [], data: [ [ "Frumpo McNugget", "12345" ] ] };
    spreadsheetTypes.Each( ( type )=>{
			var path = variables[ "temp" & type & "Path" ];
			s.newChainable( type )
        .addRows( expected.data )
				.write( path, true );
			var actual = s.read( src=path, format="array");
      expect( actual ).toBe( expected );
    })
  })

  it( "Uses the specified header row for column names", ()=>{
    var expected = { columns: [ "name", "number" ], data: [ [ "Frumpo McNugget", "12345" ] ] };
    spreadsheetTypes.Each( ( type )=>{
			var path = variables[ "temp" & type & "Path" ];
			s.newChainable( type )
        .addRow( expected.columns )
        .addRows( expected.data )
				.write( path, true );
			var actual = s.read( src=path, format="array", headerRow=1 );
      expect( actual ).toBe( expected );
    })
  })

  it( "Returns an empty array if format=array and there are no visible sheets", ()=>{
    var expected = [ columns: [], data: [] ];
		spreadsheetTypes.Each( ( type )=>{
			var path = variables[ "temp" & type & "Path" ];
			s.newChainable( type )
				.renameSheet( "hidden sheet", 1 )
				.setCellValue( "I'm in a hidden sheet", 1, 1 )
				.hideSheet( sheetNumber=1 )
			  .write( path, true );
			var actual = s.read( src=path, format="array" );
			expect( actual ).toBe( expected );
		})
	})

  it( "Reads from the specified sheet name or number", ()=>{
		var expected = { columns: [], data: [ [ "Frumpo McNugget", "12345" ] ] };
    spreadsheetTypes.Each( ( type )=>{
			var path = variables[ "temp" & type & "Path" ];
			s.newChainable( type )
        .createSheet( "sheet2" )
        .setActiveSheet( "sheet2" )
        .addRows( expected.data )
				.write( path, true );
			var actual = s.read( src=path, format="array", sheetName="sheet2" );
      expect( actual ).toBe( expected );
      actual = s.read( src=path, format="array", sheetNumber=2 );
      expect( actual ).toBe( expected );
    })
	})

  it( "Generates default column names if the data has more columns than the specifed header row", ()=>{
		var headerRow = [ "firstColumn" ];
		var dataRow1 = [ "row 1 col 1 value" ];
		var dataRow2 = [ "row 2 col 1 value", "row 2 col 2 value" ];
		var expected = { columns: [ "firstColumn", "column2" ], data: [ dataRow1, dataRow2 ] };
    spreadsheetTypes.Each( ( type )=>{
			var path = variables[ "temp" & type & "Path" ];
      s.newChainable( type )
        .addRow( headerRow )
        .addRow( dataRow1 )
        .addRow( dataRow2 )
        .write( path, true );
      var actual = s.read( src=path, format="array", headerRow=1 );
      expect( actual ).toBe( expected );
    })
	})

  it( "Includes the specified header row in array if includeHeader is true", ()=>{
    var columns = [ "name", "number" ];
    var data = [ "Frumpo McNugget", "12345" ];
    var expected = { columns: columns, data: [ columns, data ] };
    spreadsheetTypes.Each( ( type )=>{
      var path = variables[ "temp" & type & "Path" ];
      s.newChainable( type )
        .addRow( columns )
        .addRow( data )
        .write( path, true );
      var actual = s.read( src=path, format="array", headerRow=1, includeHeaderRow=true );
      expect( actual ).toBe( expected );
    })
	})

  it( "Excludes null and blank rows in array by default", ()=>{
    var data = [ [ "", "" ], [ "a", "b" ] ];
    var expected = { columns: [], data: [ [ "a", "b" ] ] };
    spreadsheetTypes.Each( ( type )=>{
      var path = variables[ "temp" & type & "Path" ];
      s.newChainable( type )
        .addRows( data )
        .write( path, true );
      var actual = s.read( src=path, format="array" );
      expect( actual ).toBe( expected );
    })
	})

  it( "Includes null and blank rows in array if includeBlankRows is true", ()=>{
    var data = [ [ "", "" ], [ "a", "b" ] ];
    var expected = { columns: [], data: data };
    spreadsheetTypes.Each( ( type )=>{
      var path = variables[ "temp" & type & "Path" ];
      s.newChainable( type )
        .addRows( data )
        .write( path, true );
      var actual = s.read( src=path, format="array", includeBlankRows=true );
      expect( actual ).toBe( expected );
    })
	})

  it( "Can handle null/empty cells", ()=>{
		var path = getTestFilePath( "nullCell.xls" );
		var actual = s.read( src=path, format="array", headerRow=1 );
		var expected = { columns: [ "column1", "column2" ] , data: [ [ "", "a" ] ] };
		expect( actual ).toBe( expected );
	})

  it( "Includes trailing empty columns when using a header row", ()=>{
		var columns = [ "column1", "emptyColumn" ];
    var data = [ "column 1 value" ];
    var expected = { columns: columns, data: [ data ] };
    spreadsheetTypes.Each( ( type )=>{
      var path = variables[ "temp" & type & "Path" ];
      s.newChainable( type )
				.addRow( columns )
				.addRow( data )
				.write( path, true );
      var actual = s.read( src=path, format="array", headerRow=1 );
      expect( actual ).toBe( expected );
    })
	})

  it( "Reads values of different types correctly, by default returning the raw values", ()=>{
		var numericValue = 2;
		var dateValue = CreateDate( 2015, 04, 12 );
		var rawDecimalValue = 0.000011;
		var leadingZeroValue = "01";
		var columns = [ "numeric", "zero", "decimal", "boolean", "date", "leadingZero" ];
		var data = [ [ numericValue, 0, rawDecimalValue, true, dateValue, leadingZeroValue ] ];
    var expected = { columns: columns, data: data };
		variables.spreadsheetTypes.Each( ( type )=>{
			var path = variables[ "temp" & type & "Path" ];
			s.newChainable( type )
				.addRows( data )
				.formatCell( { dataformat: "0.00000" }, 1, 3 )
				.write( path, true );
			var actual = s.read( src=path, format="array", columnNames=columns );
			expect( actual ).toBe( expected );
	  })
	})

  it( "Can return the visible/formatted value rather than raw value", ()=>{
		var rawDecimalValue = 0.000011;
		var visibleDecimalValue = 0.00001;
		var data = [ [ rawDecimalValue ] ];
    var expected = { columns: [], data: [ [ visibleDecimalValue ] ] };
		variables.spreadsheetTypes.Each( ( type )=>{
			var path = variables[ "temp" & type & "Path" ];
			s.newChainable( type )
				.addRows( data )
				.formatCell( { dataformat: "0.00000" }, 1, 1 )
				.write( path, true );
			var actual = s.read( src=path, format="array", returnVisibleValues=true );
			expect( actual ).toBe( expected );
			var decimalHasBeenOutputInScientificNotation = ( Trim( actual.data[ 1 ][ 1 ] ).FindNoCase( "E" ) > 0 );
			expect( decimalHasBeenOutputInScientificNotation ).toBeFalse();
		})
	})

  it( "Can read specified rows only into a query", ()=>{
    variables.spreadsheetTypes.Each( ( type )=>{
			var path = variables[ "temp" & type & "Path" ];
      var data = [ [ "row1" ], [ "row2" ] ];
      s.newChainable( type )
				.addRows( data )
				.write( path, true );
      var actual = s.read( src=path, format="array", rows="2" );
      var expected = { columns: [], data: [ [ "row2" ] ] };
      expect( actual ).toBe( expected );
    })
  })

  it( "Can read specified columns only into a query", ()=>{
    variables.spreadsheetTypes.Each( ( type )=>{
			var path = variables[ "temp" & type & "Path" ];
      var data = [ "column1", "column2" ];
      s.newChainable( type )
				.addRow( data )
				.write( path, true );
      var actual = s.read( src=path, format="array", columns="2" );
      var expected = { columns: [], data: [ [ "column2" ] ] };
      expect( actual ).toBe( expected );
    })
  })

  it( "Returns column names but no data if column names are specified but spreadsheet is empty", ()=>{
    variables.spreadsheetTypes.Each( ( type )=>{
			var path = variables[ "temp" & type & "Path" ];
      s.newChainable( type ).write( path, true );
      var actual = s.read( src=path, format="array", columnNames="One,Two" );
			var expected = { columns: [ "One", "Two" ], data: [] };
			expect( actual ).toBe( expected );
		})
	})

})
</cfscript>