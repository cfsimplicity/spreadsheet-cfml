component extends="base"{

	void function downloadBinaryVariable( required binaryVariable, required string filename, required contentType ){
		cfheader( name="Content-Disposition", value='attachment; filename="#arguments.filename#"' );
		cfcontent( type=arguments.contentType, variable="#arguments.binaryVariable#", reset="true" );
	}

	any function encryptFile( required string filepath, required string password, required string algorithm ){
		// See https://poi.apache.org/encryption.html
		// NB: Not all spreadsheet programs support this type of encryption
		// set up the encryptor with the chosen algo
		var validAlgorithms = [ "agile", "standard", "binaryRC4" ];
		if( !ArrayFindNoCase( validAlgorithms, arguments.algorithm ) )
			Throw( type=library().getExceptionType() & ".invalidAlgorithm", message="Invalid algorithm", detail="'#arguments.algorithm#' is not a valid algorithm. Supported algorithms are: #validAlgorithms.ToList( ', ')#" );
		lock name="#arguments.filepath#" timeout=5 {
			var mode = library().createJavaObject( "org.apache.poi.poifs.crypt.EncryptionMode" );
			var info = library().createJavaObject( "org.apache.poi.poifs.crypt.EncryptionInfo" ).init( mode[ arguments.algorithm ] );
			var encryptor = info.getEncryptor();
			encryptor.confirmPassword( JavaCast( "string", arguments.password ) );
			try{
				// set up a POI filesystem object
				var poifs = library().createJavaObject( "org.apache.poi.poifs.filesystem.POIFSFileSystem" );
				try{
					// set up an encrypted stream within the POI filesystem
					// ACF gets confused by encryptor.getDataStream( POIFSFileSystem ) signature. Using getRoot() means getDataStream( DirectoryNode ) will be used
					if( library().getIsACF() )
						var encryptedStream = encryptor.getDataStream( poifs.getRoot() );
					else
						var encryptedStream = encryptor.getDataStream( poifs );
					// read in the unencrypted wb file and write it to the encrypted stream
					var workbook = getWorkbookHelper().workbookFromFile( arguments.filepath );
					workbook.write( encryptedStream );
				}
				finally{
					// make sure encrypted stream in closed
					closeLocalFileOrStream( local, "encryptedStream" );
				}
				try{
					// write the encrypted POI filesystem to file, replacing the unencypted version
					var outputStream = CreateObject( "java", "java.io.FileOutputStream" ).init( arguments.filepath );
					poifs.writeFilesystem( outputStream );
					outputStream.flush();
				}
				finally{
					// always close the stream. otherwise file may be left in a locked state if an unexpected error occurs
					closeLocalFileOrStream( local, "outputStream" );
				}
			}
			finally{
				closeLocalFileOrStream( local, "poifs" );
			}
		}
		return this;
	}

	any function closeLocalFileOrStream( required localScope, required string varName ){
		if( arguments.localScope.KeyExists( arguments.varName ) )
			arguments.localScope[ arguments.varName ].close();
		return this;
	}

	string function filenameSafe( required string input ){
		var charsToRemove	=	"\|\\\*\/\:""<>~&";
		var result = arguments.input.reReplace( "[#charsToRemove#]+", "", "ALL" ).Left( 255 );
		if( result.IsEmpty() )
			return "renamed"; // in case all chars have been replaced (unlikely but possible)
		return result;
	}

	string function getFileContentTypeFromPath( required string path ){
		try{
			return FileGetMimeType( arguments.path, true ).ListLast( "/" );
		}
		catch( any exception ){
			return "unknown";
		}
	}

	void function handleInvalidSpreadsheetFile( required string path ){
		var detail = "The file #arguments.path# does not appear to be a binary or xml spreadsheet.";
		if( isCsvTsvOrTextFile( arguments.path ) )
			detail &= " It may be a CSV/TSV file, in which case use 'readCSV()' to read it";
		Throw( type="cfsimplicity.spreadsheet.invalidFile", message="Invalid spreadsheet file", detail=detail );
	}

	any function throwErrorIFfileNotExists( required string path ){
		if( !FileExists( arguments.path ) )
			getExceptionHelper().throwNonExistentFileException( arguments.path );
		return this;
	}

	any function throwErrorIFnotCsvOrTextFile( required string path ){
		if( !isCsvTsvOrTextFile( arguments.path ) )
			Throw( type=library().getExceptionType() & ".invalidCsvFile", message="Invalid csv file", detail="#arguments.path# does not appear to be a csv/tsv/text file" );
		return this;
	}	

	/* Private */

	private boolean function isCsvTsvOrTextFile( required string path ){
		var contentType = getFileContentTypeFromPath( arguments.path );
		return ListFindNoCase( "csv,tab-separated-values,plain", contentType );//Lucee=text/plain ACF=text/csv tsv=text/tab-separated-values
	}

}