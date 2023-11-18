component accessors="true"{

	property name="filepath";
	property name="firstRowIsHeader" type="boolean" default=false;
	property name="headerValues" type="array";
	property name="numberOfRowsToSkip" default=0;
	property name="returnFormat" default="none";
	property name="rowFilter";
	property name="rowProcessor";
	/* Java objects */
	property name="format"; //org.apache.commons.csv.CSVFormat
	/* Internal */
	property name="library" setter="false";

	public ReadCsv function init( required spreadsheetLibrary, required string filepath ){
		variables.library = arguments.spreadsheetLibrary;
		variables.library.getFileHelper()
			.throwErrorIFfileNotExists( arguments.filepath )
			.throwErrorIFnotCsvOrTextFile( arguments.filepath );
		variables.filepath = arguments.filepath;
		variables.format = createPredefinedFormat();
		return this;
	}

	/* Public builder API */

	public ReadCsv function intoAnArray(){
		variables.returnFormat = "array";
		return this;
	}

	public ReadCsv function withPredefinedFormat( required string type ){
		variables.format = createPredefinedFormat( arguments.type );
		return this;
	}

	/* Format configuration */
	public ReadCsv function withAllowMissingColumnNames( boolean state=true ){
		variables.format = variables.format.builder().setAllowMissingColumnNames( JavaCast( "boolean", arguments.state ) ).build();
		return this;
	}

	public ReadCsv function withAutoFlush( boolean state=true ){
		variables.format = variables.format.builder().setAutoFlush( JavaCast( "boolean", arguments.state ) ).build();
		return this;
	}

	public ReadCsv function withCommentMarker( required string marker ){
		variables.format = variables.format.builder().setCommentMarker( JavaCast( "char", arguments.marker ) ).build();
		return this;
	}

	public ReadCsv function withDelimiter( required string delimiter ){
		if( variables.library.getCsvHelper().delimiterIsTab( arguments.delimiter ) ){
			variables.format = createPredefinedFormat( "TDF" ); //tabs require several specific settings so use predefined format
			return this;
		}
		variables.format = variables.format.builder().setDelimiter( JavaCast( "string", arguments.delimiter ) ).build();
		return this;
	}

	public ReadCsv function withDuplicateHeaderMode( required string value ){
		var mode = variables.library.createJavaObject( "org.apache.commons.csv.DuplicateHeaderMode" )[ JavaCast( "string", arguments.value ) ];
		variables.format = variables.format.builder().setDuplicateHeaderMode( mode ).build();
		return this;
	}

	public ReadCsv function withEscapeCharacter( required string character ){
		variables.format = variables.format.builder().setEscape( JavaCast( "char", arguments.character ) ).build();
		return this;
	}

	public ReadCsv function withHeader( required array header ){
		variables.headerValues = arguments.header;
		variables.format = variables.format.builder().setHeader( JavaCast( "string[]", arguments.header ) ).build();
		return this;
	}

	public ReadCsv function withHeaderComments( required array comments ){
		variables.format = variables.format.builder().setHeaderComments( JavaCast( "string[]", arguments.comments ) ).build();
		return this;
	}

	public ReadCsv function withIgnoreEmptyLines( boolean state=true ){
		variables.format = variables.format.builder().setIgnoreEmptyLines( JavaCast( "boolean", arguments.state ) ).build();
		return this;
	}

	public ReadCsv function withIgnoreHeaderCase( boolean state=true ){
		variables.format = variables.format.builder().setIgnoreHeaderCase( JavaCast( "boolean", arguments.state ) ).build();
		return this;
	}

	public ReadCsv function withIgnoreSurroundingSpaces( boolean state=true ){
		variables.format = variables.format.builder().setIgnoreSurroundingSpaces( JavaCast( "boolean", arguments.state ) ).build();
		return this;
	}

	public ReadCsv function withNullString( required string value ){
		variables.format = variables.format.builder().setNullString( JavaCast( "string", arguments.value ) ).build();
		return this;
	}

	public ReadCsv function withQuoteCharacter( string character ){
		variables.format = variables.format.builder().setQuote( JavaCast( "char", arguments.character ) ).build();
		return this;
	}

	public ReadCsv function withSkipHeaderRecord( boolean state=true ){
		variables.format = variables.format.builder().setSkipHeaderRecord( JavaCast( "boolean", arguments.state ) ).build();
		return this;
	}

	public ReadCsv function withTrailingDelimiter( boolean state=true ){
		variables.format = variables.format.builder().setTrailingDelimiter( JavaCast( "boolean", arguments.state ) ).build();
		return this;
	}

	public ReadCsv function withTrim( boolean state=true ){
		variables.format = variables.format.builder().setTrim( JavaCast( "boolean", arguments.state ) ).build();
		return this;
	}

	// additional features
	public ReadCsv function withFirstRowIsHeader( boolean state=true ){
		variables.firstRowIsHeader = arguments.state;
		return this;
	}

	public ReadCsv function withSkipFirstRows( required numeric numberOfRowsToSkip ){
		if( !IsValid( "integer", arguments.numberOfRowsToSkip ) || ( arguments.numberOfRowsToSkip < 0 ) )
			Throw( type=variables.library.getExceptionType() & ".invalidArgument", message="Invalid argument to method withSkipFirstRows()", detail="'#arguments.numberOfRowsToSkip#' is not a valid argument to withSkipFirstRows(). Please specify zero or a positive integer" );
		variables.numberOfRowsToSkip = arguments.numberOfRowsToSkip;
		return this;
	}

	public ReadCsv function withRowFilter( required function rowFilter ){
		variables.rowFilter = arguments.rowFilter;
		return this;
	}

	public ReadCsv function withRowProcessor( required function rowProcessor ){
		variables.rowProcessor = arguments.rowProcessor;
		return this;
	}

	// final execution
	public any function execute(){
		if( variables.returnFormat == "array" )
			var result = [ columns: [], data: [] ];//ordered struct
		var skippedRecords = 0;
		var currentRecordNumber = 0;
		try {
			var parser = variables.library.createJavaObject( "org.apache.commons.csv.CSVParser" )
				.parse(
					CreateObject( "java", "java.io.File" ).init( variables.filepath )
					,CreateObject( "java", "java.nio.charset.Charset" ).forName( "UTF-8" )
					,variables.format
				);
			var recordIterator = parser.iterator();
			while( recordIterator.hasNext() ) {
				var values = recordIterator.next().values();
				if( skipThisRecord( skippedRecords ) ){
					skippedRecords++;
					continue;
				}
				if( variables.firstRowIsHeader && IsNull( variables.headerValues ) ){
					variables.headerValues = values;
					result.columns = values;
					continue;
				}
				if( !IsNull( variables.rowFilter ) && !variables.rowFilter( values ) )
					continue;
				if( !IsNull( variables.rowProcessor ) )
					values = variables.rowProcessor( values, ++currentRecordNumber );
				if( variables.returnFormat == "array" )
					result.data.Append( values );
			}
		}
		finally {
			variables.library.getFileHelper().closeLocalFileOrStream( local, "parser" );
		}
		if( variables.returnFormat == "array" ){
			useManuallySpecifiedHeaderForColumnsIfRequired( result );
			return result;
		}
		return this;
	}

	/* Private */
	private void function useManuallySpecifiedHeaderForColumnsIfRequired( required struct result ){
		if( ArrayLen( arguments.result.columns ) || IsNull( variables.format.getHeader() ) )
			return;
		arguments.result.columns = variables.format.getHeader();
	}

	private boolean function skipThisRecord( required numeric skippedRecords ){
		return variables.numberOfRowsToSkip && ( arguments.skippedRecords < variables.numberOfRowsToSkip );
	}

	private any function createPredefinedFormat( string type="DEFAULT" ){
		return variables.library.createJavaObject( "org.apache.commons.csv.CSVFormat" )[ JavaCast( "string", arguments.type ) ];
	}

}