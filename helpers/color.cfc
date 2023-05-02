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
		return arguments.cellFont.getHSSFColor( arguments.workbook )?.getTriplet()?:[];
	}

	any function getColor( required workbook, required string colorValue ){
		/*
			if colorValue is a preset name, returns the index
			if colorValue is hex it will be converted to RGB
			if colorValue is an RGB Triplet eg. "255,255,255" then the exact color object is returned for xlsx, or the nearest color's index if xls
		*/
		var isRGB = ListLen( arguments.colorValue ) == 3;
		if( !isRGB && !isHexColor( arguments.colorValue ) )
			return getColorIndex( arguments.colorValue );
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
			JavaCast( "int", arguments.rgb[ 1 ] )
			,JavaCast( "int", arguments.rgb[ 2 ] )
			,JavaCast( "int", arguments.rgb[ 3 ] )
		];
		try{
			return library().createJavaObject( "org.apache.poi.xssf.usermodel.XSSFColor" ).init( JavaCast( "byte[]", rgbBytes ), JavaCast( "null", 0 ) );
		}
		//ACF doesn't handle signed java byte values the same way as Lucee: see https://www.bennadel.com/blog/2689-creating-signed-java-byte-values-using-coldfusion-numbers.htm
		catch( any exception ){
			if( !exception.message CONTAINS "cannot fit inside a byte" )
				rethrow;
			//ACF2016+ Bitwise operators can't handle >32-bit args: https://stackoverflow.com/questions/43176313/cffunction-cfargument-pass-unsigned-int32
			var javaLangInteger = CreateObject( "java", "java.lang.Integer" );
			var negativeMask = InputBaseN( ( "11111111" & "11111111" & "11111111" & "00000000" ), 2 );
			negativeMask = javaLangInteger.parseUnsignedInt( negativeMask );
			rgbBytes = [];
			for( var value in arguments.rgb ){
				if( BitMaskRead( value, 7, 1 ) )
				value = BitOr( negativeMask, value );//value greater than 127
				rgbBytes.Append( JavaCast( "byte", value ) );
			}
			return library().createJavaObject( "org.apache.poi.xssf.usermodel.XSSFColor" ).init( JavaCast( "byte[]", rgbBytes ), JavaCast( "null", 0 ) );
		}
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

	private numeric function getColorIndex( required string colorName ){
		var findColor = arguments.colorName.Trim().UCase();
		//check for 9 extra colours from old org.apache.poi.ss.usermodel.IndexedColors and map
		var deprecatedNames = [ "BLACK1", "WHITE1", "RED1", "BRIGHT_GREEN1", "BLUE1", "YELLOW1", "PINK1", "TURQUOISE1", "LIGHT_TURQUOISE1" ];
		if( ArrayFind( deprecatedNames, findColor ) )
			findColor = findColor.Left( findColor.Len() - 1 );
		var indexedColors = library().createJavaObject( "org.apache.poi.hssf.util.HSSFColor$HSSFColorPredefined" );
		try{
			var color = indexedColors.valueOf( JavaCast( "string", findColor ) );
			return color.getIndex();
		}
		catch( any exception ){
			Throw( type=library().getExceptionType() & ".invalidColor", message="Invalid Color", detail="The color provided (#arguments.colorName#) is not valid. Use getPresetColorNames() for a list of valid color names" );
		}
	}

	private boolean function isHexColor( required string inputString ){
		return arguments.inputString.REFind( "^##?[0-9A-Fa-f]{6,6}$" );
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
		if( IsNull( arguments.cellFont.getXSSFColor()?.getRGB() ) )
			return [];
		return convertSignedRGBToPositiveTriplet( arguments.cellFont.getXSSFColor().getRGB() );
	}

}