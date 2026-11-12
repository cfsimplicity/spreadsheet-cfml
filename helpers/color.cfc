component extends="base"{

	array function convertSignedRGBToPositiveTriplet( required any signedRGB ){
		// When signed, values of 128+ are negative: convert then to positive values
		var result = [];
		for( var i=1; i <= 3; i++ ){
			result.Append( ( arguments.signedRGB[ i ] < 0 )? ( arguments.signedRGB[ i ] + 256 ): arguments.signedRGB[ i ] );
		}
		return result;
	}

	array function getRGBFromCellFont( required workbook, required any cellFont ){
		if( library().isXmlFormat( arguments.workbook ) )
			return getRGBFromXSSFCellFont( arguments.cellFont );
		var hssfColor = arguments.cellFont.getHSSFColor( arguments.workbook );
		if( IsNull( hssfColor ) )
			return [];
		return hssfColor.getTriplet();
	}

	any function getColor( required workbook, required string colorValue ){
		/*
			if colorValue is a preset name, returns the index for xls and the color object for xlsx
			if colorValue is hex it will be converted to RGB
			if colorValue is an RGB Triplet eg. "255,255,255" then the exact color object is returned for xlsx, or the nearest color's index if xls
		*/
		var isRGB = ListLen( arguments.colorValue ) == 3;
		var isPresetColor = !isRGB && !isHexColor( arguments.colorValue );
		if( isPresetColor ){
			var presetColor = getPresetColor( arguments.colorValue );
			if( library().isBinaryFormat( arguments.workbook ) )
				return presetColor.getIndex();
			return getColorForXlsx( presetColor.getTriplet() );
		}
		if( !isRGB && isHexColor( arguments.colorValue ) )
			arguments.colorValue = hexToRGB( arguments.colorValue );
		var rgb = ListToArray( arguments.colorValue );
		if( library().isXmlFormat( arguments.workbook ) )
			return getColorForXlsx( rgb );
		var palette = arguments.workbook.getCustomPalette();
		var similarExistingColor = palette.findSimilarColor(
			JavaCast( "int", rgb[ 1 ] )
			,JavaCast( "int", rgb[ 2 ] )
			,JavaCast( "int", rgb[ 3 ] )
		);
		return similarExistingColor.getIndex();
	}

	struct function getJavaColorRGBFor( required string colorName ){
		var findColor = arguments.colorName.Trim().UCase();
		var color = CreateObject( "Java", "java.awt.Color" );
		if( IsNull( color[ findColor ] ) || !IsInstanceOf( color[ findColor ], "java.awt.Color" ) )//don't use member functions on color
			Throw( type=library().getExceptionType() & ".invalidColor", message="Invalid color", detail="The color provided (#arguments.colorName#) is not valid." );
		color = color[ findColor ];
		var colorRGB = {
			red: color.getRed()
			,green: color.getGreen()
			,blue: color.getBlue()
		};
		return colorRGB;
	}

	string function getRgbTripletForStyleColorFormat( required workbook, required cellStyle, required string format ){
		var isXlsx = library().isXmlFormat( arguments.workbook );
		var palette = isXlsx? "": arguments.workbook.getCustomPalette();
		var colorObject = getColorObjectForFormat( arguments.format, arguments.cellStyle, palette, isXlsx );
		// HSSF will return an empty string rather than a null if the color doesn't exist
		if( IsNull( colorObject ) || IsSimpleValue( colorObject) )
			return "";
		return getRGBStringFromColorObject( colorObject );
	}

	string function getRGBStringFromColorObject( required colorObject ){
		if( arguments.colorObject.getClass().getSimpleName() == "XSSFColor" )
			return ArrayToList( convertSignedRGBToPositiveTriplet( arguments.colorObject.getRGB() ) );
		return ArrayToList( arguments.colorObject.getTriplet() );
	}
	
	/* Private */
	private any function getColorForXlsx( required array rgb ){
		var rgbBytes = [
			CreateObject( "java", "java.lang.Integer" ).valueOf( arguments.rgb[ 1 ] ).byteValue()
			,CreateObject( "java", "java.lang.Integer" ).valueOf( arguments.rgb[ 2 ] ).byteValue()
			,CreateObject( "java", "java.lang.Integer" ).valueOf( arguments.rgb[ 3 ] ).byteValue()
		];
		return library().createJavaObject( "org.apache.poi.xssf.usermodel.XSSFColor" ).init( JavaCast( "byte[]", rgbBytes ), JavaCast( "null", 0 ) );
	}

	private any function getColorObjectForFormat(
		required string format
		,required cellStyle
		,required any palette
		,required boolean isXlsx
	){
		switch( arguments.format ){
			case "bottombordercolor":
				return arguments.isXlsx? arguments.cellStyle.getBottomBorderXSSFColor(): arguments.palette.getColor( arguments.cellStyle.getBottomBorderColor() );
			case "fgcolor":
				return arguments.isXlsx? arguments.cellStyle.getFillForegroundXSSFColor(): arguments.palette.getColor( arguments.cellStyle.getFillForegroundColor() );
			case "leftbordercolor":
				return arguments.isXlsx? arguments.cellStyle.getLeftBorderXSSFColor(): arguments.palette.getColor( arguments.cellStyle.getLeftBorderColor() );
			case "rightbordercolor":
				return arguments.isXlsx? arguments.cellStyle.getRightBorderXSSFColor(): arguments.palette.getColor( arguments.cellStyle.getRightBorderColor() );
			case "topbordercolor":
				return arguments.isXlsx? arguments.cellStyle.getTopBorderXSSFColor(): arguments.palette.getColor( arguments.cellStyle.getTopBorderColor() );
		}
	}

	private any function getPresetColor( required string colorName ){
		var findColor = arguments.colorName.Trim().UCase();
		//check for 9 extra colours from old org.apache.poi.ss.usermodel.IndexedColors and map
		var deprecatedNames = [ "BLACK1", "WHITE1", "RED1", "BRIGHT_GREEN1", "BLUE1", "YELLOW1", "PINK1", "TURQUOISE1", "LIGHT_TURQUOISE1" ];
		if( ArrayFind( deprecatedNames, findColor ) )
			findColor = findColor.Left( findColor.Len() - 1 );
		var presetColors = library().createJavaObject( "org.apache.poi.hssf.util.HSSFColor$HSSFColorPredefined" );
		try{
			return presetColors.valueOf( JavaCast( "string", findColor ) );
		}
		catch( any exception ){
			Throw( type=library().getExceptionType() & ".invalidColor", message="Invalid Color", detail="The color provided (#arguments.colorName#) is not valid. Use getPresetColorNames() for a list of valid color names" );
		}
	}

	private boolean function isHexColor( required string inputString ){
		return JavaCast( "String", arguments.inputString ).matches( "^##?[0-9A-Fa-f]{6,6}$" );
	}

	private string function hexToRGB( required string hexColor ){
		if( !isHexColor( arguments.hexColor ) )
			return "";
		arguments.hexColor = arguments.hexColor.Replace( "##", "" );
		var response = [];
		for( var i=1; i <= 5; i=i+2 ){
			response.Append( InputBaseN( Mid( arguments.hexColor, i, 2 ), 16 ) );
		}
		return response.ToList();
	}

	private array function getRGBFromXSSFCellFont( required any cellFont ){
		var xssfColor = arguments.cellFont.getXSSFColor();
		if( IsNull( xssfColor ) )
			return [];
		var rgb = xssfColor.getRGB();
		if( IsNull( rgb ) )
			return [];
		return convertSignedRGBToPositiveTriplet( rgb );
	}

}