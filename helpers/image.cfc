component extends="base" accessors="true"{

	numeric function addImageToWorkbook(
		required workbook
		,required any image //path or object
		,string imageType
	){
		// TODO image objects don't always work, depending on how they're created: POI accepts it but the image is not displayed (broken)
		var imageArgumentIsObject = IsImage( arguments.image );
		if( imageArgumentIsObject && !arguments.KeyExists( "imageType" ) )
			Throw( type=library().getExceptionType(), message="Invalid argument combination", detail="If you specify an image object, you must also provide the imageType argument" );
		var imageArgumentIsFile = ( !imageArgumentIsObject && IsSimpleValue( arguments.image ) && FileExists( arguments.image ) );
		if( !imageArgumentIsObject && IsSimpleValue( arguments.image ) && !imageArgumentIsFile )
			getExceptionHelper().throwNonExistentFileException( arguments.image );
		if( !imageArgumentIsObject && !imageArgumentIsFile )
			Throw( type=library().getExceptionType(), message="Invalid image", detail="You must provide either a file path or an image object" );
		if( imageArgumentIsFile ){
			arguments.imageType = getFileHelper().getFileContentTypeFromPath( arguments.image );
			if( arguments.imageType == "unknown" )
				getExceptionHelper().throwUnknownImageTypeException();
		}
		var imageTypeIndex = getImageTypeIndex( arguments.workbook, arguments.imageType );
		var bytes = imageArgumentIsFile? FileReadBinary( arguments.image ): ToBinary( ToBase64( arguments.image ) );
		return arguments.workbook.addPicture( bytes, JavaCast( "int", imageTypeIndex ) );// returns 1-based integer index
	}

	/* Private */

	private numeric function getImageTypeIndex( required workbook, required string imageType ){
		switch( arguments.imageType ){
			case "DIB": case "EMF": case "JPEG": case "PICT": case "PNG": case "WMF":
				return arguments.workbook[ "PICTURE_TYPE_" & arguments.imageType.UCase() ];
			case "JPG":
				return arguments.workbook.PICTURE_TYPE_JPEG;
		}
		Throw( type=library().getExceptionType(), message="Invalid Image Type", detail="Valid image types are DIB, EMF, JPG, JPEG, PICT, PNG, and WMF" );
	}

}