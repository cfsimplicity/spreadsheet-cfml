<cfscript>
describe( "processLargeFile", ()=>{

  it( "can read and process an xlsx file one row at a time using a passed UDF", ()=>{
    s.newChainable( "xlsx" ).addRows( [ [ 1, 2 ], [ 3, 4 ] ] ).write( tempXlsxPath, true );
    variables.tempTotal = 0;
    var sumOfValues = 10;
    var processor = ( rowValues )=>{
      rowValues.Each( ( value )=>{
        tempTotal = ( tempTotal + value );
      })
    };
    s.processLargeFile( tempXlsxPath )
      .withRowProcessor( processor )
      .execute();
    expect( tempTotal ).toBe( sumOfValues );
  })

  it( "passes the current record number to the processor UDF", ()=>{
    s.newChainable( "xlsx" ).addRows( [ [ 1, 2 ], [ 3, 4 ] ] ).write( tempXlsxPath, true );
    var result = [];
    var processor = ( rowValues, rowNumber )=> result.Append( rowNumber );
    s.processLargeFile( tempXlsxPath )
      .withRowProcessor( processor )
      .execute();
    var expected = [ 1, 2 ];
    expect( result ).toBe( expected );
  })

  it( "passes column names/headers to the processor UDF", ()=>{
    s.newChainable( "xlsx" )
      .addRows( [ [ "first", "last" ], [ "Frumpo", "McNugget" ], [ "Susi", "Sorglos" ] ] )
      .write( tempXlsxPath, true );
    var result = [];
    var processor = ( rowValues, rowNumber, columnNames )=>{
      var row = {};
      ArrayEach( columnNames, ( columnName, index )=>{
        row[ columnName ] = rowValues[ index ]
      });
      result.Append( row.first );
    };
    s.processLargeFile( tempXlsxPath )
      .withRowProcessor( processor )
      .withFirstRowIsHeader()
      .execute();
    var expected = [ "Frumpo", "Susi" ];
    expect( result ).toBe( expected );
  })

  it( "allows streaming reader tuning options to be set", ()=> {
    var options = {
      bufferSize: 2048
      ,rowCacheSize: 20
    };
    var processObject = s.processLargeFile( tempXlsxPath ).withStreamingOptions( options );
    expect( processObject.getStreamingOptions() ).toBe( options );
  })

  it( "can process an encrypted XLSX file", ()=>{
		var path = getTestFilePath( "passworded.xlsx" );
    var result = "";
    var processor = ( rowValues )=> { result = rowValues };
    s.processLargeFile( path )
      .withPassword( "pass" )
      .withRowProcessor( processor )
      .execute();
    var expected = [ "secret" ];
		expect( result ).toBe( expected );
	})

  it( "processes the specified sheet number", ()=>{
		s.newChainable( "xlsx" )
      .createSheet( "SecondSheet" )
      .setActiveSheet( "SecondSheet" )
      .addRow( [ "test" ] )
      .write( tempXlsxPath, true );
    var result = [];
    var processor = ( rowValues )=> { result = rowValues };
    x = s.processLargeFile( tempXlsxPath )
      .withSheetNumber( 2 )
      .withRowProcessor( processor )
      .execute();
    var expected = [ "test" ];
		expect( result ).toBe( expected );
	})

  it( "processes the specified sheet name", ()=>{
		s.newChainable( "xlsx" )
      .createSheet( "SecondSheet" )
      .setActiveSheet( "SecondSheet" )
      .addRow( [ "test" ] )
      .write( tempXlsxPath, true );
    var result = [];
    var processor = ( rowValues )=>{ result = rowValues; };
    s.processLargeFile( tempXlsxPath )
      .withSheetName( "SecondSheet" )
      .withRowProcessor( processor )
      .execute();
    var expected = [ "test" ];
		expect( result ).toBe( expected );
	})

  it( "can process visible/formatted values rather than raw values", ()=>{
		var rawValue = 0.000011;
		var visibleValue = 0.00001;
		s.newChainable( "xlsx" )
			.setCellValue( rawValue, 1, 1, "numeric" )
			.formatCell( { dataformat: "0.00000" }, 1, 1 )
			.write( tempXlsxPath, true );
		var result = [];
    var processor = ( rowValues )=>{ result = rowValues; };
    s.processLargeFile( tempXlsxPath )
      .withRowProcessor( processor )
      .withUseVisibleValues( true )
      .execute();
		expect( result[ 1 ] ).toBe( visibleValue );
	})

  it( "can skip the first N rows", ()=> {
    s.newChainable( "xlsx" )
      .addRows( [ [ "skip me" ], [ "skip me too" ], [ "data" ] ] )
      .write( tempXlsxPath, true );
    var result = [];
    var processor = ( rowValues )=> result.Append( rowValues );
    s.processLargeFile( tempXlsxPath )
      .withRowProcessor( processor )
      .withSkipFirstRows( 2 )
      .execute();
    var expected = [ [ "data" ] ];
    expect( result ).toBe( expected );
  })

  it( "can ignore the first row if it contains the headers", ()=> {
    s.newChainable( "xlsx" )
      .addRows( [ [ "heading" ], [ "data" ] ] )
      .write( tempXlsxPath, true );
    var result = [];
    var processor = ( rowValues )=> result.Append( rowValues );
    s.processLargeFile( tempXlsxPath )
      .withRowProcessor( processor )
      .withFirstRowIsHeader()
      .execute();
    var expected = [ [ "data" ] ];
    expect( result ).toBe( expected );
  })

  it( "will treat the first non-skipped row as the header if both options specified", ()=> {
    s.newChainable( "xlsx" )
      .addRows( [ [ "skip me" ], [ "header" ], [ "data" ] ] )
      .write( tempXlsxPath, true );
    var result = [];
    var processor = ( rowValues )=> result.Append( rowValues );
    s.processLargeFile( tempXlsxPath )
      .withRowProcessor( processor )
      .withFirstRowIsHeader()
      .withSkipFirstRows( 1 )
      .execute();
    var expected = [ [ "data" ] ];
    expect( result ).toBe( expected );
  })

  describe( "processLargeFile throws an exception if", ()=>{

		it( "the file doesn't exist", ()=>{
			expect( ()=>{
				var path = getTestFilePath( "nonexistent.xls" );
				s.processLargeFile( path ).execute();
			}).toThrow( type="cfsimplicity.spreadsheet.nonExistentFile" );
		})
		
		it( "the file to be read is not an XLSX type", ()=>{
			expect( ()=>{
				var path = getTestFilePath( "test.xls" );
				s.processLargeFile( path ).execute();
			}).toThrow( type="cfsimplicity.spreadsheet.invalidSpreadsheetType" );
		})

    it( "the source file is not a spreadsheet", ()=>{
			expect( ()=>{
				s.processLargeFile( getTestFilePath( "notaspreadsheet.txt" ) ).execute();
			}).toThrow( type="cfsimplicity.spreadsheet.invalidSpreadsheetType" );
		})

		it( "the sheet name doesn't exist", ()=>{
			expect( ()=>{
        s.processLargeFile( getTestFilePath( "large.xlsx" ) ).withSheetName( "nonexistent" ).execute();
			}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetName" );
		})

		it( "the sheet number doesn't exist", ()=>{
			expect( ()=>{
				s.processLargeFile( getTestFilePath( "large.xlsx" ) ).withSheetNumber( 20 ).execute();
			}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetNumber" );
		})

	})
  
})
</cfscript>