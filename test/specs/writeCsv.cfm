<cfscript>
describe( "writeCsv", ()=>{

	//Note: a trailing newline is always expected when printing from Commons CSV

	it( "writeCsv defaults to the EXCEL predefined format", ()=>{
		var object = s.writeCsv();
		var format = object.getFormat();
		expect( format.equals( format.EXCEL ) ).toBeTrue();
	})

	describe( "writeCsv can write a csv file or return a csv string", ()=>{

		afterEach( ()=>{
			if( FileExists( tempCsvPath ) )
				FileDelete( tempCsvPath );
		})

		it( "from an array of arrays", ()=>{
			var data = [ [ "a", "b" ], [ "c", "d" ] ];
			var expected = "a,b#newline#c,d#newline#";
			var actual = s.writeCsv()
				.fromData( data )
				.execute();
			expect( actual ).toBe( expected );
			s.writeCsv()
				.toFile( tempCsvPath )
				.fromData( data )
				.execute();
			actual = FileRead( tempCsvPath );
			expect( actual ).toBe( expected );
		})

		it( "from an array of structs", ()=>{
			var data = [ [ first: "Frumpo", last: "McNugget" ] ];
			var expected = "Frumpo,McNugget#newline#";
			var actual = s.writeCsv()
				.fromData( data )
				.execute();
			expect( actual ).toBe( expected );
			s.writeCsv()
				.toFile( tempCsvPath )
				.fromData( data )
				.execute();
			actual = FileRead( tempCsvPath );
			expect( actual ).toBe( expected );
		})

		it( "from a query", ()=>{
			var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
			var expected = "a,b#newline#c,d#newline#";
			var actual = s.writeCsv()
				.fromData( data )
				.execute();
			expect( actual ).toBe( expected );
			s.writeCsv()
				.toFile( tempCsvPath )
				.fromData( data )
				.execute();
			actual = FileRead( tempCsvPath );
			expect( actual ).toBe( expected );
		})

	})

	it( "allows an alternative to the default comma delimiter", ()=>{
		var data = [ [ "a", "b" ], [ "c", "d" ] ];
		var expected = "a|b#newline#c|d#newline#";
		var actual = s.writeCsv()
			.fromData( data )
			.withDelimiter( "|" )
			.execute();
		expect( actual ).toBe( expected );
	})

	it( "has special handling when specifying tab as the delimiter", ()=>{
		var data = [ [ "a", "b" ], [ "c", "d" ] ];
		var validTabValues = [ "#Chr( 9 )#", "\t", "tab", "TAB" ];
		var expected = "a#Chr( 9 )#b#newline#c#Chr( 9 )#d#newline#";
		for( var delimiter in validTabValues ){
			var actual = s.writeCsv()
				.fromData( data )
				.withDelimiter( delimiter )
				.execute();
			expect( actual ).toBe( expected );
		}
	})

	it( "can use the query columns as the header row", ()=>{
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
		var expected = "column1,column2#newline#a,b#newline#c,d#newline#";
		var actual = s.writeCsv()
			.fromData( data )
			.withQueryColumnsAsHeader()
			.execute();
		expect( actual ).toBe( expected );
	})

	it( "can use the row struct keys as the header row", ()=>{
		var data = [ [ first: "Frumpo", last: "McNugget" ] ];
		var expected = "first,last#newline#Frumpo,McNugget#newline#";
		var actual = s.writeCsv()
			.fromData( data )
			.withStructKeysAsHeader()
			.execute();
		expect( actual ).toBe( expected );
	})

	it( "outputs integers correctly with no decimal point", ()=>{
		var arrayData = [ [ 123 ] ];
		var queryData = QueryNew( "column1", "Integer", arrayData );
		var expected = "123#newline#";
		expect( s.writeCsv().fromData( arrayData ).execute() ).toBe( expected );
		expect( s.writeCsv().fromData( queryData ).execute() ).toBe( expected );
	})

	it( "outputs date objects using the instance's specified DATETIME format", ()=>{
		var nowAsText = DateTimeFormat( Now(), s.getDateFormats().DATETIME );
		var arrayData = [ [ ParseDateTime( nowAsText ) ] ];
		var queryData = QueryNew( "column1", "Timestamp", arrayData );
		var expected = "#nowAsText##newline#";
		expect( s.writeCsv().fromData( arrayData ).execute() ).toBe( expected );
		expect( s.writeCsv().fromData( queryData ).execute() ).toBe( expected );
	})

	it( "does NOT treat date strings as date objects to be formatted using the DATETIME format", ()=>{
		var dateString = "2022-12-18";
		var data = [ [ dateString ] ];
		var expected = '#dateString##newline#';
		expect( s.writeCsv().fromData( data ).execute() ).toBe( expected );
	})

	it( "can handle an embedded delimiter", ()=>{
		var data = [ [ "a,a", "b" ], [ "c", "d" ] ];
		var expected = '"a,a",b#newline#c,d#newline#';
		expect( s.writeCsv().fromData( data ).execute() ).toBe( expected );
	})

	it( "can handle an embedded double-quote", ()=>{
		var data = [ [ "a""a", "b" ], [ "c", "d" ] ];
		var expected = '"a""a",b#newline#c,d#newline#';
		expect( s.writeCsv().fromData( data ).execute() ).toBe( expected );
	})

	it( "can handle an embedded carriage return", ()=>{
		var data = [ [ "a#newline#a", "b" ], [ "c", "d" ] ];
		var expected = '"a#newline#a",b#newline#c,d#newline#';
		expect( s.writeCsv().fromData( data ).execute() ).toBe( expected );
	})

	it(
		title="can process rows in parallel if the engine supports it"
		,body=function(){
			//can't test if using threads, just that there are no errors
			var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "a" ], [ "a", "a" ] ] );
			var expected = "a,a#newline#a,a#newline#";//same values because order is not guaranteed
			var actual = s.writeCsv()
				.fromData( data )
				.withParallelThreads( 2 )
				.execute();
			expect( actual ).toBe( expected );
		}
		//20231031: ACF 2021 and 2023 won't run the whole suite if this test is included: testbox errors thrown
		//running just the queryToCsv tests works fine though. Lucee is fine with the whole suite.
		,skip=s.getIsACF()
	);

	it( "allows Commons CSV format options to be applied", ()=>{
		var path = getTestFilePath( "test.csv" );
		var object = s.writeCsv()
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
			.withQuoteMode( "NON_NUMERIC" )
			.withSkipHeaderRecord()
			.withTrailingDelimiter()
			.withTrim();
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
		expect( object.getFormat().getQuoteMode().name() ).toBe( "NON_NUMERIC" );
		expect( object.getFormat().getSkipHeaderRecord() ).toBeTrue();
		expect( object.getFormat().getTrailingDelimiter() ).toBeTrue();
		expect( object.getFormat().getTrim() ).toBeTrue();
		//reverse check in case any of the above were defaults
		object
			.withAutoFlush( false )
			.withDuplicateHeaderMode( "ALLOW_ALL" )
			.withIgnoreEmptyLines( false )
			.withIgnoreHeaderCase( false )
			.withIgnoreSurroundingSpaces( false )
			.withQuoteMode( "MINIMAL" )
			.withSkipHeaderRecord( false )
			.withTrailingDelimiter( false )
			.withTrim( false );
		expect( object.getFormat().getAutoFlush() ).toBeFalse();
		expect( object.getFormat().getDuplicateHeaderMode().name() ).toBe( "ALLOW_ALL" );
		expect( object.getFormat().getIgnoreEmptyLines() ).toBeFalse();
		expect( object.getFormat().getIgnoreHeaderCase() ).toBeFalse();
		expect( object.getFormat().getIgnoreSurroundingSpaces() ).toBeFalse();
		expect( object.getFormat().getQuoteMode().name() ).toBe( "MINIMAL" );
		expect( object.getFormat().getSkipHeaderRecord() ).toBeFalse();
		expect( object.getFormat().getTrailingDelimiter() ).toBeFalse();
		expect( object.getFormat().getTrim() ).toBeFalse();
	})

	describe( "writeCsv() throws an exception if", ()=>{

		it( "executed with no data", ()=>{
			expect( ()=>{
				s.writeCsv().execute();
			}).toThrow( type="cfsimplicity.spreadsheet.missingDataForCsv" );
		})

		it( "the data is not an array or query", ()=>{
			expect( ()=>{
				var data = "string";
				s.writeCsv().fromData( data ).execute();
			}).toThrow( type="cfsimplicity.spreadsheet.invalidDataForCsv" );
		})

		it( "the data contains complex values", ()=>{
			expect( ()=>{
				var complexValue = [];
				var data = [ [ complexValue ] ];
				s.writeCsv().fromData( data ).execute();
			}).toThrow( type="cfsimplicity.spreadsheet.invalidDataForCsv" );
		})

		it(
			title="parallel threads are specified and the engine does not support it"
			,body=function(){
				expect( ()=>{
					s.writeCsv().withParallelThreads();
				}).toThrow( type="cfsimplicity.spreadsheet.parallelOptionNotSupported" );
			}
			,skip=function(){
				return s.engineSupportsParallelLoopProcessing();
			}
		);

		it( "the file path specified is VFS", ()=>{
			expect( ()=>{
				var data = [ [ "a", "b" ], [ "c", "d" ] ];
				var path = "ram://temp.csv";
				s.writeCsv().fromData( data ).toFile( path ).execute();
			}).toThrow( type="cfsimplicity.spreadsheet.vfsNotSupported" );
		})

	})

})
</cfscript>