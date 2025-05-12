component extends="BaseCsv" accessors="true"{

	property name="firstRowIsHeader" type="boolean" default="false";
	property name="numberOfRowsToSkip" default=0;
	property name="processRowsAsJavaArrays" type="boolean" default="true";
	property name="returnFormat" default="none";
	property name="rowFilter";
	property name="rowProcessor";

	public ReadCsv function init( required spreadsheetLibrary, required string filepath ){
		super.init( arguments.spreadsheetLibrary );
		variables.library.getFileHelper()
			.throwErrorIFfileNotExists( arguments.filepath )
			.throwErrorIFnotCsvOrTextFile( arguments.filepath );
		variables.filepath = arguments.filepath;
		return this;
	}

	/* Public builder API */

	public ReadCsv function intoAnArray(){
		variables.returnFormat = "array";
		return this;
	}

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

	public ReadCsv function processRowsAsJavaArrays( boolean state=true ){
		variables.processRowsAsJavaArrays = arguments.state;
		return this;
	}

	// final execution
	public any function execute(){
		var result = [ columns: [], data: [] ];//ordered struct
		var skippedRecords = 0;
		var currentRecordNumber = 0;
		try {
			// Remove any Byte Order Mark from beginning of CSV
			var BOMInputStream = variables.library.createJavaObject( "org.apache.commons.io.input.BOMInputStream" )
				.builder()
				.setPath( variables.filepath )
				.get();
			var inputStreamReader = CreateObject( "java", "java.io.InputStreamReader" ).init( BOMInputStream, "UTF-8" );
			var parser = variables.format.builder()
				.get()
				.parse( inputStreamReader );
			var recordIterator = parser.iterator();
			while( recordIterator.hasNext() ) {
				var values = recordIterator.next().values();
				if( skipThisRecord( skippedRecords ) ){
					skippedRecords++;
					continue;
				}
				if( !variables.processRowsAsJavaArrays )
					values = convertJavaArrayToCFML( values );
				if( variables.firstRowIsHeader && IsNull( variables.headerValues ) ){
					variables.headerValues = values;
					result.columns = values;
					continue;
				}
				if( !IsNull( variables.rowFilter ) && !variables.rowFilter( values, result.columns ) )
					continue;
				if( !IsNull( variables.rowProcessor ) )
					values = variables.rowProcessor( values, ++currentRecordNumber, result.columns );
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

	private function convertJavaArrayToCFML( required javaArray ){
		return ArrayNew( 1 ).Append( arguments.javaArray, true );
	}

}