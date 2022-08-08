component extends="base" accessors="true"{

	/* Common exceptions */
	void function throwOldExcelFormatException( required string path ){
		Throw( type="#library().getExceptionType()#.OldExcelFormatException", message="Invalid spreadsheet format", detail="The file #arguments.path# was saved in a format that is too old. Please save it as an 'Excel 97/2000/XP' file or later." );
	}

	void function throwFileExistsException( required string path ){
		Throw( type=library().getExceptionType(), message="File already exists", detail="The file path #arguments.path# already exists. Use 'overwrite=true' if you wish to overwrite it." );
	}

	void function throwNonExistentFileException( required string path ){
		Throw( type=library().getExceptionType(), message="Non-existent file", detail="Cannot find the file #arguments.path#." );
	}

	void function throwUnknownImageTypeException(){
		Throw( type=library().getExceptionType(), message="Could not determine image type", detail="An image type could not be determined from the image provided" );
	}

	void function throwExceptionIFreadFormatIsInvalid(){
		if( arguments.KeyExists( "format" ) && !ListFindNoCase( "query,html,csv", arguments.format ) )
			Throw( type=library().getExceptionType() & ".invalidReadFormat", message="Invalid format", detail="Supported formats are: 'query', 'html' and 'csv'" );
	}

	void function throwInvalidFileForReadLargeFileException(){
		Throw( type=library().getExceptionType() & ".invalidFile", message="Invalid spreadsheet file", detail="readLargeFile() can only be used with XLSX files. The file you are trying to read does not appear to be an XLSX file." );
	}

}