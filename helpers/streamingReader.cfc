component extends="base"{

	boolean function isStreamingReaderFormat( required workbook ){
		return arguments.workbook.getClass().getCanonicalName() == "com.github.pjfanning.xlsx.impl.StreamingWorkbook";
	}

	query function readFileIntoQuery( required string path, required struct builderOptions, required struct sheetToQueryArgs ){
		lock name="#arguments.path#" timeout=5 {
			try{
				var file = CreateObject( "java", "java.io.FileInputStream" ).init( arguments.path );
				arguments.sheetToQueryArgs.workbook = getBuilder( arguments.builderOptions ).open( file );
				return getSheetHelper().sheetToQuery( argumentCollection=arguments.sheetToQueryArgs );
			}
			catch( any exception ){
				getExceptionHelper().throwExceptionIfFileIsInvalidForStreamingReader( exception );
				rethrow;
			}
			finally{
				getFileHelper().closeLocalFileOrStream( local, "file" );
				getFileHelper().closeLocalFileOrStream( local, "workbook" );
			}
		}
	}

	// NB: called from tests
	any function getBuilder( required struct options ){
		var passwordProtected = ( arguments.options.KeyExists( "password") && arguments.options.password.Trim().Len() );
		var builder = library().createJavaObject( "com.github.pjfanning.xlsx.StreamingReader" ).builder()
			.setFullFormatRichText( JavaCast( "boolean", true ) ); //some sheet methods e.g. getLastRowNum() may error if not set to true!
		if( passwordProtected )
			builder.password( JavaCast( "string", arguments.options.password ) );
		if( arguments.options.KeyExists( "bufferSize" ) )
			builder.bufferSize( JavaCast( "int", arguments.options.bufferSize ) );
		if( arguments.options.KeyExists( "rowCacheSize" ) )
			builder.rowCacheSize( JavaCast( "int", arguments.options.rowCacheSize ) );
		return builder;
	}

}