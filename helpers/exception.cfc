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

}