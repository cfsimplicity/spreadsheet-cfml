component extends="base" accessors="true"{

	public any function createWorkBook(
		required string sheetName
		,boolean xmlFormat=false
		,boolean streamingXml=false
		,numeric streamingWindowSize=100
	){
		getSheetHelper().validateSheetName( arguments.sheetName );
		if( !arguments.xmlFormat )
			return getClassHelper().loadClass( library().getHSSFWorkbookClassName() ).init();
		if( !arguments.streamingXml )
			return getClassHelper().loadClass( library().getXSSFWorkbookClassName() ).init();
		if( !IsValid( "integer", arguments.streamingWindowSize ) || ( arguments.streamingWindowSize < 1 ) )
			Throw( type=library().getExceptionType(), message="Invalid 'streamingWindowSize' argument", detail="'streamingWindowSize' must be an integer value greater than 1" );
		return getClassHelper().loadClass( library().getSXSSFWorkbookClassName() ).init( JavaCast( "int", arguments.streamingWindowSize ) );
	}

	public any function workbookFromFile( required string path, string password ){
		// works with both xls and xlsx
		// see https://stackoverflow.com/a/46149469 for why FileInputStream is preferable to File
		// 20210322 using File doesn't seem to improve memory usage anyway.
		lock name="#arguments.path#" timeout=5 {
			try{
				var factory = getClassHelper().loadClass( "org.apache.poi.ss.usermodel.WorkbookFactory" );
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

}