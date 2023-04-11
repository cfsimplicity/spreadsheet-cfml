component extends="base"{

	any function createWorkBook( string type, numeric streamingWindowSize=100 ){
		if( arguments.type == "xls" )
			return library().createJavaObject( getClassHelper().getClassName( "HSSFWorkbook" ) ).init();
		if( arguments.type == "xlsx" )
			return library().createJavaObject( getClassHelper().getClassName( "XSSFWorkbook" ) ).init();
		// Streaming Xlsx
		if( !IsValid( "integer", arguments.streamingWindowSize ) || ( arguments.streamingWindowSize < 1 ) )
			Throw( type=library().getExceptionType() & ".invalidStreamingWindowSizeArgument", message="Invalid 'streamingWindowSize' argument", detail="'streamingWindowSize' must be an integer value greater than 1" );
		return library().createJavaObject( getClassHelper().getClassName( "SXSSFWorkbook" ) ).init( JavaCast( "int", arguments.streamingWindowSize ) );
	}

	any function workbookFromFile( required string path, string password ){
		// works with both xls and xlsx
		// see https://stackoverflow.com/a/46149469 for why FileInputStream is preferable to File
		// 20210322 using File doesn't seem to improve memory usage anyway.
		lock name="#arguments.path#" timeout=5 {
			try{
				var factory = library().createJavaObject( "org.apache.poi.ss.usermodel.WorkbookFactory" );
				var file = CreateObject( "java", "java.io.FileInputStream" ).init( arguments.path );
				if( arguments.KeyExists( "password" ) )
					return factory.create( file, arguments.password );
				return factory.create( file );
			}
			catch( org.apache.poi.hssf.OldExcelFormatException exception ){
				getExceptionHelper().throwOldExcelFormatException( arguments.path );
			}
			catch( any exception ){
				if( exception.message CONTAINS "unsupported file type" )
					getFileHelper().handleInvalidSpreadsheetFile( arguments.path );// from POI 5.x
				if( exception.message CONTAINS "spreadsheet seems to be Excel 5" )
					getExceptionHelper().throwOldExcelFormatException( arguments.path );
				rethrow;
			}
			finally{
				getFileHelper().closeLocalFileOrStream( local, "file" );
			}
		}
	}

	string function typeFromArguments( boolean xmlFormat=false, boolean streamingXml=false ){
		if( !arguments.xmlFormat && !arguments.streamingXml )
			return "xls";
		if( arguments.streamingXml )
			return "streamingXlsx";
		return "xlsx";
	}

}