component extends="base"{

	numeric function addImageToWorkbook(
		required workbook
		,required any image //path or object
		,string imageType
	){
		// TODO image objects don't always work, depending on how they're created: POI accepts it but the image is not displayed (broken)
		var imageArgumentIsObject = IsImage( arguments.image );
		if( imageArgumentIsObject && !arguments.KeyExists( "imageType" ) )
			Throw( type=library().getExceptionType() & ".invalidArgumentCombination", message="Invalid argument combination", detail="If you specify an image object, you must also provide the imageType argument" );
		var imageArgumentIsFile = ( !imageArgumentIsObject && IsSimpleValue( arguments.image ) && FileExists( arguments.image ) );
		if( !imageArgumentIsObject && IsSimpleValue( arguments.image ) && !imageArgumentIsFile )
			getExceptionHelper().throwNonExistentFileException( arguments.image );
		if( !imageArgumentIsObject && !imageArgumentIsFile )
			Throw( type=library().getExceptionType() & ".invalidImage", message="Invalid image", detail="You must provide either a file path or an image object" );
		if( imageArgumentIsFile ){
			arguments.imageType = getFileHelper().getFileContentTypeFromPath( arguments.image );
			if( arguments.imageType == "unknown" )
				getExceptionHelper().throwUnknownImageTypeException();
		}
		var imageTypeIndex = getImageTypeIndex( arguments.workbook, arguments.imageType );
		var bytes = imageArgumentIsFile? FileReadBinary( arguments.image ): ToBinary( ToBase64( arguments.image ) );
		return arguments.workbook.addPicture( bytes, JavaCast( "int", imageTypeIndex ) );// returns 1-based integer index
	}

	any function createAnchor( required workbook, required array anchorCoordinates ){
		var clientAnchorClass = library().isXmlFormat( arguments.workbook )
				? "org.apache.poi.xssf.usermodel.XSSFClientAnchor"
				: "org.apache.poi.hssf.usermodel.HSSFClientAnchor";
		var anchor = library().createJavaObject( clientAnchorClass ).init();
		if( arguments.anchorCoordinates.Len() == 4 ){
			anchor.setRow1( JavaCast( "int", arguments.anchorCoordinates[ 1 ] -1 ) );
			anchor.setCol1( JavaCast( "int", arguments.anchorCoordinates[ 2 ] -1 ) );
			anchor.setRow2( JavaCast( "int", arguments.anchorCoordinates[ 3 ] -1 ) );
			anchor.setCol2( JavaCast( "int", arguments.anchorCoordinates[ 4 ] -1 ) );
			return anchor;
		}
		anchor.setDx1( JavaCast( "int", arguments.anchorCoordinates[ 1 ] ) );
		anchor.setDy1( JavaCast( "int", arguments.anchorCoordinates[ 2 ] ) );
		anchor.setDx2( JavaCast( "int", arguments.anchorCoordinates[ 3 ] ) );
		anchor.setDy2( JavaCast( "int", arguments.anchorCoordinates[ 4 ] ) );
		anchor.setRow1( JavaCast( "int", arguments.anchorCoordinates[ 5 ] -1 ) );
		anchor.setCol1( JavaCast( "int", arguments.anchorCoordinates[ 6 ] -1 ) );
		anchor.setRow2( JavaCast( "int", arguments.anchorCoordinates[ 7 ] -1 ) );
		anchor.setCol2( JavaCast( "int", arguments.anchorCoordinates[ 8 ] -1 ) );
		return anchor;
	}

	/* Private */

	private numeric function getImageTypeIndex( required workbook, required string imageType ){
		switch( arguments.imageType ){
			case "DIB": case "EMF": case "JPEG": case "PICT": case "PNG": case "WMF":
				return arguments.workbook[ "PICTURE_TYPE_" & arguments.imageType.UCase() ];
			case "JPG":
				return arguments.workbook.PICTURE_TYPE_JPEG;
		}
		Throw( type=library().getExceptionType() & ".invalidImageType", message="Invalid Image Type", detail="Valid image types are DIB, EMF, JPG, JPEG, PICT, PNG, and WMF" );
	}

}