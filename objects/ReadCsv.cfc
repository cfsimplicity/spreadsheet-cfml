component extends="BaseCsv" accessors="true"{

	property name="firstRowIsHeader" type="boolean" default="false";
	property name="maxNumberOfColumns" type="integer" default=0;
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

	public ReadCsv function intoAQuery(){
		variables.returnFormat = "query";
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
				if( !variables.processRowsAsJavaArrays || ( variables.returnFormat == "query" ) )
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
				if( ( variables.returnFormat == "array" ) || ( variables.returnFormat == "query" ) )
					result.data.Append( values );

				if( variables.returnFormat == "query" )
					variables.maxNumberOfColumns = Max( ArrayLen( values ), variables.maxNumberOfColumns );//query conversion requires consistency between headers/data
			}
		}
		finally {
			variables.library.getFileHelper().closeLocalFileOrStream( local, "parser" );
		}
		if( variables.returnFormat == "array" ){
			useManuallySpecifiedHeaderForColumnsIfRequired( result );
			return result;
		}
		if( variables.returnFormat == "query" ){
			useManuallySpecifiedHeaderForColumnsIfRequired( result );
			return convertResultToQuery( result );
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

	private query function convertResultToQuery( required struct result ){
		var numberOfHeaders = ArrayLen( arguments.result.columns );
		throwErrorIfHeaderColumnMismatch( numberOfHeaders );
		if( numberOfHeaders == 0 )
			generateDefaultColumns( arguments.result );
		equalizeColumnLengths( arguments.result );
		return DeserializeJson( SerializeJson( arguments.result ), false );
	}

	private void function generateDefaultColumns( required struct result ){
		for( var i=1; i <= variables.maxNumberOfColumns; i++ )
			ArrayAppend( arguments.result.columns, "column" & i );
	}

	private void function equalizeColumnLengths( required struct result ){
		arguments.result.data.Each( function( row, index ){
			ArrayResize( result.data[ index ], variables.maxNumberOfColumns );//don't scope arguments within closure
		});
	}

	private void function throwErrorIfHeaderColumnMismatch( required numeric numberOfHeaders ){
		if( arguments.numberOfHeaders == 0 )
			return;
		if( arguments.numberOfHeaders != variables.maxNumberOfColumns )
 			Throw( type=variables.library.getExceptionType() & ".invalidCsvHeaders", message="Invalid CSV headers/column names", detail="The number of headers (#arguments.numberOfHeaders#) doesn't match the number of columns (#variables.maxNumberOfColumns#)" );
	}

}