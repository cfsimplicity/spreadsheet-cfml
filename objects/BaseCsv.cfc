component accessors="true"{

	property name="filepath";
	property name="headerValues" type="array";
	/* Java objects */
	property name="format"; //org.apache.commons.csv.CSVFormat
	/* Internal */
	property name="library" setter="false";

	public BaseCsv function init( required spreadsheetLibrary, string initialPredefinedFormat="DEFAULT" ){
		variables.library = arguments.spreadsheetLibrary;
		variables.format = createPredefinedFormat( arguments.initialPredefinedFormat );
		return this;
	}

	/* Public builder API */
	public BaseCsv function withPredefinedFormat( required string type ){
		variables.format = createPredefinedFormat( arguments.type );
		return this;
	}

	/* Format configuration */
	public BaseCsv function withAllowMissingColumnNames( boolean state=true ){
		variables.format = variables.format.builder().setAllowMissingColumnNames( JavaCast( "boolean", arguments.state ) ).build();
		return this;
	}

	public BaseCsv function withAutoFlush( boolean state=true ){
		variables.format = variables.format.builder().setAutoFlush( JavaCast( "boolean", arguments.state ) ).build();
		return this;
	}

	public BaseCsv function withCommentMarker( required string marker ){
		variables.format = variables.format.builder().setCommentMarker( JavaCast( "char", arguments.marker ) ).build();
		return this;
	}

	public BaseCsv function withDelimiter( required string delimiter ){
		if( variables.library.getCsvHelper().delimiterIsTab( arguments.delimiter ) ){
			variables.format = createPredefinedFormat( "TDF" ); //tabs require several specific settings so use predefined format
			return this;
		}
		variables.format = variables.format.builder().setDelimiter( JavaCast( "string", arguments.delimiter ) ).build();
		return this;
	}

	public BaseCsv function withDuplicateHeaderMode( required string value ){
		var mode = variables.library.createJavaObject( "org.apache.commons.csv.DuplicateHeaderMode" )[ JavaCast( "string", arguments.value ) ];
		variables.format = variables.format.builder().setDuplicateHeaderMode( mode ).build();
		return this;
	}

	public BaseCsv function withEscapeCharacter( required string character ){
		variables.format = variables.format.builder().setEscape( JavaCast( "char", arguments.character ) ).build();
		return this;
	}

	public BaseCsv function withHeader( required array header ){
		variables.headerValues = arguments.header;
		variables.format = variables.format.builder().setHeader( JavaCast( "string[]", arguments.header ) ).build();
		return this;
	}

	public BaseCsv function withHeaderComments( required array comments ){
		variables.format = variables.format.builder().setHeaderComments( JavaCast( "string[]", arguments.comments ) ).build();
		return this;
	}

	public BaseCsv function withIgnoreEmptyLines( boolean state=true ){
		variables.format = variables.format.builder().setIgnoreEmptyLines( JavaCast( "boolean", arguments.state ) ).build();
		return this;
	}

	public BaseCsv function withIgnoreHeaderCase( boolean state=true ){
		variables.format = variables.format.builder().setIgnoreHeaderCase( JavaCast( "boolean", arguments.state ) ).build();
		return this;
	}

	public BaseCsv function withIgnoreSurroundingSpaces( boolean state=true ){
		variables.format = variables.format.builder().setIgnoreSurroundingSpaces( JavaCast( "boolean", arguments.state ) ).build();
		return this;
	}

	public BaseCsv function withNullString( required string value ){
		variables.format = variables.format.builder().setNullString( JavaCast( "string", arguments.value ) ).build();
		return this;
	}

	public BaseCsv function withQuoteCharacter( string character ){
		variables.format = variables.format.builder().setQuote( JavaCast( "char", arguments.character ) ).build();
		return this;
	}

	public BaseCsv function withQuoteMode( required string value ){
		var mode = variables.library.createJavaObject( "org.apache.commons.csv.QuoteMode" )[ JavaCast( "string", arguments.value ) ];
		variables.format = variables.format.builder().setQuoteMode( mode ).build();
		return this;
	}

	public BaseCsv function withSkipHeaderRecord( boolean state=true ){
		variables.format = variables.format.builder().setSkipHeaderRecord( JavaCast( "boolean", arguments.state ) ).build();
		return this;
	}

	public BaseCsv function withTrailingDelimiter( boolean state=true ){
		variables.format = variables.format.builder().setTrailingDelimiter( JavaCast( "boolean", arguments.state ) ).build();
		return this;
	}

	public BaseCsv function withTrim( boolean state=true ){
		variables.format = variables.format.builder().setTrim( JavaCast( "boolean", arguments.state ) ).build();
		return this;
	}

	//Private
	private any function createPredefinedFormat( string type="DEFAULT" ){
		return variables.library.createJavaObject( "org.apache.commons.csv.CSVFormat" )[ JavaCast( "string", arguments.type ) ];
	}

}