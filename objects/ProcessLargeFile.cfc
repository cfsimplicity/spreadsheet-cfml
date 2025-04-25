component accessors="true"{

  property name="filepath";
	property name="firstRowIsHeader" type="boolean" default="false";
	property name="numberOfRowsToSkip" default=0;
  property name="rowProcessor";
	property name="sheetName";
	property name="sheetNumber" default=1;
  property name="streamingOptions";
	property name="useVisibleValues" type="boolean" default="false";
  /* Internal */
	property name="library" setter="false";

  public ProcessLargeFile function init( required spreadsheetLibrary, required string filepath ){
		variables.library = arguments.spreadsheetLibrary;
		variables.library.getFileHelper().throwErrorIFfileNotExists( arguments.filepath );
		variables.filepath = arguments.filepath;
    variables.streamingOptions = {};
		return this;
	}

  /* Public builder API */

	public ProcessLargeFile function withFirstRowIsHeader( boolean state=true ){
		variables.firstRowIsHeader = arguments.state;
		return this;
	}

  public ProcessLargeFile function withRowProcessor( required function rowProcessor ){
		variables.rowProcessor = arguments.rowProcessor;
		return this;
	}

	public ProcessLargeFile function withPassword( required string password ){
		variables.streamingOptions.password = arguments.password;
		return this;
	}

	public ProcessLargeFile function withSheetName( required string sheetName ){
		variables.sheetName = arguments.sheetName;
		return this;
	}

	public ProcessLargeFile function withSheetNumber( required numeric sheetNumber ){
		variables.sheetNumber = arguments.sheetNumber;
		return this;
	}

	public ProcessLargeFile function withSkipFirstRows( required numeric numberOfRowsToSkip ){
		if( !IsValid( "integer", arguments.numberOfRowsToSkip ) || ( arguments.numberOfRowsToSkip < 0 ) )
			Throw( type=variables.library.getExceptionType() & ".invalidArgument", message="Invalid argument to method withSkipFirstRows()", detail="'#arguments.numberOfRowsToSkip#' is not a valid argument to withSkipFirstRows(). Please specify zero or a positive integer" );
		variables.numberOfRowsToSkip = arguments.numberOfRowsToSkip;
		return this;
	}

	public ProcessLargeFile function withStreamingOptions( required struct options ){
		if( arguments.options.KeyExists( "bufferSize" ) )
			variables.streamingOptions.bufferSize = Val( arguments.options.bufferSize );
		if( arguments.options.KeyExists( "rowCacheSize" ) )
			variables.streamingOptions.rowCacheSize = Val( arguments.options.rowCacheSize );
		return this;
	}

	public ProcessLargeFile function withUseVisibleValues( boolean state=true ){
		variables.useVisibleValues = arguments.state;
		return this;
	}

  // final execution
	public ProcessLargeFile function execute(){
    lock name="#getFilepath()#" timeout=5 {
			try{
				var file = CreateObject( "java", "java.io.FileInputStream" ).init( getFilepath() );
				var workbook = getLibrary().getStreamingReaderHelper().getBuilder( getStreamingOptions() ).open( file );
				var rowIterator = getSheetToProcess( workbook ).rowIterator();
				var currentRecordNumber = 0;
				var columns = [];
				var headerRowSkipped = false;
				var skippedRecords = 0;
				var rowProcessor = getRowProcessor()
				var rowDataArgs = {
					workbook: workbook
				};
				if( getUseVisibleValues() )
					rowDataArgs.returnVisibleValues = true;
				while( rowIterator.hasNext() ){
					rowDataArgs.row = rowIterator.next();
					if( skipThisRecord( skippedRecords ) ){
						skippedRecords++;
						continue;
					}
					var data = getLibrary().getRowHelper().getRowData( argumentCollection=rowDataArgs );
					if( getFirstRowIsHeader() && !headerRowSkipped ){
						headerRowSkipped = true;
						columns	=	data; 
						continue;
					}
					if( !IsNull( rowProcessor ) )
						rowProcessor( data, ++currentRecordNumber, columns );
				}
			}
			catch( any exception ){
				getLibrary().getExceptionHelper().throwExceptionIfFileIsInvalidForStreamingReader( exception );
				rethrow;
			}
			finally{
				getLibrary().getFileHelper().closeLocalFileOrStream( local, "file" );
				getLibrary().getFileHelper().closeLocalFileOrStream( local, "workbook" );
			}
		}
		return this;
	}

	/* Private */
	private any function getSheetToProcess( required workbook ){
		if( !IsNull( getSheetName() ) )
			return getLibrary().getSheetHelper().getSheetByName( arguments.workbook, getSheetName() );
		return getLibrary().getSheetHelper().getSheetByNumber( arguments.workbook, getSheetNumber() );
	}

	private boolean function skipThisRecord( required numeric skippedRecords ){
		return variables.numberOfRowsToSkip && ( arguments.skippedRecords < variables.numberOfRowsToSkip );
	}

}