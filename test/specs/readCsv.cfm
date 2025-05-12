<cfscript>
describe( "readCsv", ()=>{

	it( "can read a basic csv file into an array", ()=>{
		var path = getTestFilePath( "test.csv" );
		var expected = { columns: [], data: [ [ "Frumpo McNugget", "12345" ] ] };
		var actual = s.readCsv( path )
			.intoAnArray()
			.execute();
		expect( actual ).toBe( expected );
	})

	it( "processes each row as a java array by default", ()=>{
		var path = getTestFilePath( "test.csv" );
		var result = s.readCsv( path )
			.intoAnArray()
			.execute();
		expect( result.data[ 1 ].getClass().getCanonicalName() ).toBe( "java.lang.String[]" );
	})

	it( "can process each row as a cfml array", ()=>{
		var path = getTestFilePath( "test.csv" );
		var result = s.readCsv( path )
			.intoAnArray()
			.processRowsAsJavaArrays( false )
			.execute();
		expect( result.data[ 1 ].getClass().getCanonicalName() ).notToBe( "java.lang.String[]" );
	})

	it( "allows predefined formats to be specified", ()=>{
		var csv = '"Frumpo McNugget"#Chr( 9 )#12345';
		FileWrite( tempCsvPath, csv );
		var expected = { columns: [], data: [ [ "Frumpo McNugget", "12345" ] ] };
		var actual = s.readCsv( tempCsvPath )
			.intoAnArray()
			.withPredefinedFormat( "TDF" )
			.execute();
		expect( actual ).toBe( expected );
	})

	it( "has special handling when specifying tab as the delimiter", ()=>{
		var csv = '"Frumpo McNugget"#Chr( 9 )#12345';
		FileWrite( tempCsvPath, csv );
		var validTabValues = [ "#Chr( 9 )#", "\t", "tab", "TAB" ];
		var expected = { columns: [], data: [ [ "Frumpo McNugget", "12345" ] ] };
		for( var delimiter in validTabValues ){
			var actual = s.readCsv( tempCsvPath )
				.intoAnArray()
				.withDelimiter( delimiter )
				.execute();
			expect( actual ).toBe( expected );
		}
	})

	it( "allows N rows to be skipped at the start of the file", ()=>{
		var csv = 'Skip this line#newline#skip this line too#newline#"Frumpo McNugget",12345';
		var expected = { columns: [], data: [ [ "Frumpo McNugget", "12345" ] ] };
		FileWrite( tempCsvPath, csv );
		var actual = s.readCsv( tempCsvPath )
			.intoAnArray()
			.withSkipFirstRows( 2 )
			.execute();
		expect( actual ).toBe( expected );
	})

	it( "removes any BOM so it is not parsed with the first value", ()=>{
		var path = getTestFilePath( "csvWithBom.csv" );
		var expected = { columns: [ "Region" ], data: [ [ "UK" ] ] };
		var actual = s.readCsv( path )
			.intoAnArray()
			.withFirstRowIsHeader()
			.execute();
		expect( actual ).toBe( expected );
		expect( actual.columns[ 1 ] == "Region" ).toBeTrue();//comparison will fail if file has BOM
	})

	describe( "auto header/column handling", ()=>{

		it( "can auto extract the column names from first row if specified", ()=>{
			var csv = 'name,number#newline#"Frumpo McNugget",12345';
			var expected = { columns: [ "name", "number" ], data: [ [ "Frumpo McNugget", "12345" ] ] };
			FileWrite( tempCsvPath, csv );
			var actual = s.readCsv( tempCsvPath )
				.intoAnArray()
				.withFirstRowIsHeader( true )
				.execute();
			expect( actual ).toBe( expected );
		})

		it( "auto extraction treats the first non-skipped row as the header", ()=>{
			var csv = 'Skip this line#newline#name,number#newline#"Frumpo McNugget",12345';
			var expected = { columns: [ "name", "number" ], data: [ [ "Frumpo McNugget", "12345" ] ] };
			FileWrite( tempCsvPath, csv );
			var actual = s.readCsv( tempCsvPath )
				.intoAnArray()
				.withSkipFirstRows( 1 )
				.withFirstRowIsHeader( true )
				.execute();
			expect( actual ).toBe( expected );
		})

		it( "adds a manually specified header row to the columns result", ()=>{
			var csv = 'name,number#newline#"Frumpo McNugget",12345';
			var expected = { columns: [ "name", "number" ], data: [ [ "Frumpo McNugget", "12345" ] ] };
			FileWrite( tempCsvPath, csv );
			var actual = s.readCsv( tempCsvPath )
				.intoAnArray()
				.withHeader( [ "name", "number" ] )
				.withSkipHeaderRecord( true )
				.execute();
			expect( actual ).toBe( expected );
		})

	})

	describe( "passing UDFs to readCsv", ()=>{

		it( "allows rows to be filtered out of processing using a passed filter UDF", ()=>{
			var csv = '"Frumpo McNugget",12345#newline#"Skip",12345#newline#"Susi Sorglos",67890';
			var expected = { columns: [], data: [ [ "Frumpo McNugget", "12345" ], [ "Susi Sorglos", "67890" ] ] };
			FileWrite( tempCsvPath, csv );
			var filter = ( rowValues )=>{
				return !ArrayFindNoCase( rowValues, "skip" );
			};
			var actual = s.readCsv( tempCsvPath )
				.intoAnArray()
				.withRowFilter( filter )
				.execute();
			expect( actual ).toBe( expected );
		})

		it( "allows rows to be processed using a passed UDF and the processed values returned", ()=>{
			var csv = '"Frumpo McNugget",12345#newline#"Susi Sorglos",67890';
			var expected = { columns: [], data: [ [ "XFrumpo McNugget", "X12345" ], [ "XSusi Sorglos", "X67890" ] ] };
			FileWrite( tempCsvPath, csv );
			var processor = ( rowValues )=>{
				//NB: rowValues is a java native array. Array member functions won't work therefore.
				return ArrayMap( rowValues, ( value )=>{
					return "X" & value;
				})
			};
			var actual = s.readCsv( tempCsvPath )
				.intoAnArray()
				.withRowProcessor( processor )
				.execute();
			expect( actual ).toBe( expected );
		})

		it( "allows rows to be processed using a passed UDF without returning any data", ()=>{
			var csv = '10#newline#10';
			var expected = 20;
			FileWrite( tempCsvPath, csv );
			variables.tempTotal = 0;
			var processor = ( rowValues )=>{
				ArrayEach( rowValues, ( value )=>{
					tempTotal = ( tempTotal + value );
				})
			};
			s.readCsv( tempCsvPath )
				.withRowProcessor( processor )
				.execute();
			expect( tempTotal ).toBe( expected );
		})

		it( "passes the current record number to the processor UDF", ()=>{
			var csv = '"Frumpo McNugget",12345#newline#"Susi Sorglos",67890';
			var expected = [ 1, 2 ];
			FileWrite( tempCsvPath, csv );
			variables.temp = [];
			var processor = ( rowValues, rowNumber )=>{
				temp.Append( rowNumber );
			};
			s.readCsv( tempCsvPath )
				.withRowProcessor( processor )
				.execute();
			expect( temp ).toBe( expected );
		})

		it( "passes column names/headers to the processor UDF", ()=>{
			var csv = 'name,number#newline#"Frumpo McNugget",12345#newline#"Susi Sorglos",67890';
			var expected = [ "Frumpo McNugget", "Susi Sorglos" ];
			FileWrite( tempCsvPath, csv );
			variables.temp = [];
			var processor = ( rowValues, rowNumber, columns )=>{
				var row = {};
				ArrayEach( columns, ( column, index )=>{
					row[ column ] = rowValues[ index ]
				});
				temp.Append( row.name );
			};
			s.readCsv( tempCsvPath )
				.withFirstRowIsHeader()
				.withRowProcessor( processor )
				.execute();
			expect( temp ).toBe( expected );
		})

		it( "passes column names/headers to the filter UDF", ()=>{
			var csv = 'name,number#newline#"Frumpo McNugget",12345#newline#"Susi Sorglos",67890';
			var expected = { columns: [ "name", "number" ], data: [ [ "Susi Sorglos", "67890" ] ] };
			FileWrite( tempCsvPath, csv );
			variables.temp = [];
			var filter = ( rowValues, columns )=>{
				var row = {};
				ArrayEach( columns, ( column, index )=>{
					row[ column ] = rowValues[ index ]
				});
				return !FindNoCase( "Frumpo", row.name );
			};
			var actual = s.readCsv( tempCsvPath )
				.intoAnArray()
				.withFirstRowIsHeader()
				.withRowFilter( filter )
				.execute();
			expect( actual ).toBe( expected );
		})

	})

	it( "allows Commons CSV format options to be applied", ()=>{
		var path = getTestFilePath( "test.csv" );
		var object = s.readCsv( path )
			.withAllowMissingColumnNames()
			.withAutoFlush()
			.withCommentMarker( "##" )
			.withDelimiter( "|" )
			.withDuplicateHeaderMode( "ALLOW_EMPTY" )
			.withEscapeCharacter( "\" )
			.withHeader( [ "Name", "Number" ] )
			.withHeaderComments( [ "comment1", "comment2" ] )
			.withIgnoreEmptyLines()
			.withIgnoreHeaderCase()
			.withIgnoreSurroundingSpaces()
			.withNullString( "" )
			.withQuoteCharacter( "'" )
			.withSkipHeaderRecord()
			.withTrailingDelimiter()
			.withTrim();
		expect( object.getFormat().getAllowMissingColumnNames() ).toBeTrue();
		expect( object.getFormat().getAutoFlush() ).toBeTrue();
		expect( object.getFormat().getCommentMarker() ).toBe( "##" );
		expect( object.getFormat().getDelimiterString() ).toBe( "|" );
		expect( object.getFormat().getDuplicateHeaderMode().name() ).toBe( "ALLOW_EMPTY" );
		expect( object.getFormat().getEscapeCharacter() ).toBe( "\" );
		expect( object.getFormat().getHeader() ).toBe( [ "Name", "Number" ] );
		expect( object.getFormat().getHeaderComments() ).toBe( [ "comment1", "comment2" ] );
		expect( object.getFormat().getIgnoreEmptyLines() ).toBeTrue();
		expect( object.getFormat().getIgnoreHeaderCase() ).toBeTrue();
		expect( object.getFormat().getIgnoreSurroundingSpaces() ).toBeTrue();
		expect( object.getFormat().getNullString() ).toBe( "" );
		expect( object.getFormat().getQuoteCharacter() ).toBe( "'" );
		expect( object.getFormat().getSkipHeaderRecord() ).toBeTrue();
		expect( object.getFormat().getTrailingDelimiter() ).toBeTrue();
		expect( object.getFormat().getTrim() ).toBeTrue();
		//reverse check in case any of the above were defaults
		object
			.withAllowMissingColumnNames( false )
			.withAutoFlush( false )
			.withDuplicateHeaderMode( "ALLOW_ALL" )
			.withIgnoreEmptyLines( false )
			.withIgnoreHeaderCase( false )
			.withIgnoreSurroundingSpaces( false )
			.withSkipHeaderRecord( false )
			.withTrailingDelimiter( false )
			.withTrim( false );
		expect( object.getFormat().getAllowMissingColumnNames() ).toBeFalse();
		expect( object.getFormat().getAutoFlush() ).toBeFalse();
		expect( object.getFormat().getDuplicateHeaderMode().name() ).toBe( "ALLOW_ALL" );
		expect( object.getFormat().getIgnoreEmptyLines() ).toBeFalse();
		expect( object.getFormat().getIgnoreHeaderCase() ).toBeFalse();
		expect( object.getFormat().getIgnoreSurroundingSpaces() ).toBeFalse();
		expect( object.getFormat().getSkipHeaderRecord() ).toBeFalse();
		expect( object.getFormat().getTrailingDelimiter() ).toBeFalse();
		expect( object.getFormat().getTrim() ).toBeFalse();
	})

	describe( "readCsv() throws an exception if", ()=>{

		it( "a zero or positive integer is not passed to withSkipFirstRows()", ()=>{
			expect( ()=>{
				var actual = s.readCsv( getTestFilePath( "test.csv" ) ).withSkipFirstRows( -1 );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidArgument" );
		})

	})

})
</cfscript>