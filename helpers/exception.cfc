component extends="base"{

	/* Common exceptions */
	void function throwOldExcelFormatException( required string path ){
		Throw( type=library().getExceptionType() & ".oldExcelFormatException", message="Invalid spreadsheet format", detail="The file #arguments.path# was saved in a format that is too old. Please save it as an 'Excel 97/2000/XP' file or later." );
	}

	void function throwFileExistsException( required string path ){
		Throw( type=library().getExceptionType() & ".fileAlreadyExists", message="File already exists", detail="The file path #arguments.path# already exists. Use 'overwrite=true' if you wish to overwrite it." );
	}

	void function throwNonExistentFileException( required string path ){
		Throw( type=library().getExceptionType() & ".nonExistentFile", message="Non-existent file", detail="Cannot find the file #arguments.path#." );
	}

	void function throwUnknownImageTypeException(){
		Throw( type=library().getExceptionType() & ".unknownImageType", message="Could not determine image type", detail="An image type could not be determined from the image provided" );
	}

	void function throwExceptionIFreadFormatIsInvalid(){
		if( arguments.KeyExists( "format" ) && !ListFindNoCase( "query,array,arrayOfStructs,html,csv", arguments.format ) )
			Throw( type=library().getExceptionType() & ".invalidReadFormat", message="Invalid format", detail="Supported formats are: 'query', 'array', 'arrayOfStructs', 'html' and 'csv'" );
	}

	void function throwExceptionIfFileIsInvalidForStreamingReader( required exception ){
		/*
			for some reason ACF won't match the exception type as a catch() arg here, i.e.
			catch( com.github.pjfanning.xlsx.exceptions.ReadException exception ){} hence using an if-test
		*/
		if( arguments.exception.type == "com.github.pjfanning.xlsx.exceptions.ReadException" )
			Throw( type=library().getExceptionType() & ".invalidSpreadsheetType", message="Invalid spreadsheet file", detail="readLargeFile() and processLargeFile() can only be used with XLSX files. The file you are trying to read does not appear to be an XLSX file." );
	}

	void function throwNonExistentRowException( required numeric rowNumber ){
		Throw( type=library().getExceptionType() & ".nonExistentRow", message="Non-existent row", detail="Row #arguments.rowNumber# doesn't exist. You may need to create the row first by adding data to it." );
	}

	void function throwParallelOptionNotSupportedException(){
		Throw( type=library().getExceptionType() & ".parallelOptionNotSupported", message="Parallel threads option not supported", detail="Your ColdFusion engine does not support parallel processing of loops via threads" );
	}

	void function throwInvalidCellException( required numeric rowNumber, required numeric columnNumber ){
		Throw( type=library().getExceptionType() & ".invalidCell", message="Invalid cell", detail="The requested cell [#arguments.rowNumber#,#arguments.columnNumber#] does not exist in the active sheet" );
	}

	void function throwFormulaEvaluationException( required cell ){
		Throw( type=library().getExceptionType() & ".failedFormula", message="Failed to run formula", detail="There is a problem with the formula in sheet #arguments.cell.getSheet().getSheetName()# row #( arguments.cell.getRowIndex() +1 )# column #( arguments.cell.getColumnIndex() +1 )#");
	}

}