component extends="base"{

	string function baseFontToHtml( required workbook, required contents, required baseFont ){
		/*
			the order of processing is important for the tests to match
			font family and size not parsed here because all cells would trigger formatting of these attributes: defaults can't be assumed
		*/
		var cssStyles = getStringHelper().newJavaStringBuilder();
		// bold
		if( arguments.baseFont.getBold() )
			cssStyles.Append( fontStyleToCss( "bold", true ) );
		// color
		if( !fontColorIsBlack( arguments.baseFont.getColor() ) )
			cssStyles.Append( fontStyleToCss( "color", arguments.baseFont.getColor(), arguments.workbook ) );
		// italic
		if( arguments.baseFont.getItalic() )
			cssStyles.Append( fontStyleToCss( "italic", true ) );
		// underline/strike
		if( arguments.baseFont.getStrikeout() || arguments.baseFont.getUnderline() ){
			var decorationValue	=	[];
			if( arguments.baseFont.getStrikeout() )
				decorationValue.Append( "line-through" );
			if( arguments.baseFont.getUnderline() )
				decorationValue.Append( "underline" );
			cssStyles.Append( fontStyleToCss( "decoration", decorationValue.ToList( " " ) ) );
		}
		cssStyles = cssStyles.toString();
		if( cssStyles.IsEmpty() )
			return arguments.contents;
		return "<span style=""#cssStyles#"">#arguments.contents#</span>";
	}

	any function cloneFont( required workbook, required fontToClone ){
		var newFont = arguments.workbook.createFont();
		// copy the existing cell's font settings to the new font
		newFont.setBold( arguments.fontToClone.getBold() );
		newFont.setCharSet( arguments.fontToClone.getCharSet() );
		// xlsx fonts contain XSSFColor objects which may have been set as RGB
		var color = library().isXmlFormat( arguments.workbook )? arguments.fontToClone.getXSSFColor(): arguments.fontToClone.getColor();
		// reportedly getXSSFColor() returns null in some conditions (not reproducible)
		if( !IsNull( color ) )
			newFont.setColor( color );
		newFont.setFontHeight( arguments.fontToClone.getFontHeight() );
		newFont.setFontName( arguments.fontToClone.getFontName() );
		newFont.setItalic( arguments.fontToClone.getItalic() );
		newFont.setStrikeout( arguments.fontToClone.getStrikeout() );
		newFont.setTypeOffset( arguments.fontToClone.getTypeOffset() );
		newFont.setUnderline( arguments.fontToClone.getUnderline() );
		return newFont;
	}

	string function runFontToHtml( required workbook, required baseFont, required runFont ){
		// NB: the order of processing is important for the tests to match
		var cssStyles = getStringHelper().newJavaStringBuilder();
		// bold
		if( Compare( arguments.runFont.getBold(), arguments.baseFont.getBold() ) )
			cssStyles.Append( fontStyleToCss( "bold", arguments.runFont.getBold() ) );
		// color
		if( Compare( arguments.runFont.getColor(), arguments.baseFont.getColor() ) && !fontColorIsBlack( arguments.runFont.getColor() ) )
			cssStyles.Append( fontStyleToCss( "color", arguments.runFont.getColor(), arguments.workbook ) );
		// italic
		if( Compare( arguments.runFont.getItalic(), arguments.baseFont.getItalic() ) )
			cssStyles.Append( fontStyleToCss( "italic", arguments.runFont.getItalic() ) );
		// underline/strike
		if( Compare( arguments.runFont.getStrikeout(), arguments.baseFont.getStrikeout() ) || Compare( arguments.runFont.getUnderline(), arguments.baseFont.getUnderline() ) ){
			var decorationValue	=	[];
			if( !arguments.baseFont.getStrikeout() && arguments.runFont.getStrikeout() )
				decorationValue.Append( "line-through" );
			if( !arguments.baseFont.getUnderline() && arguments.runFont.getUnderline() )
				decorationValue.Append( "underline" );
			//if either or both are in the base format, and either or both are NOT in the run format, set the decoration to none.
			if(
					( arguments.baseFont.getUnderline() || arguments.baseFont.getStrikeout() )
					&&
					( !arguments.runFont.getUnderline() || !arguments.runFont.getUnderline() )
				){
				cssStyles.Append( fontStyleToCss( "decoration", "none" ) );
			}
			else
				cssStyles.Append( fontStyleToCss( "decoration", decorationValue.ToList( " " ) ) );
		}
		return cssStyles.toString();
	}

	/* Private */

	private boolean function fontColorIsBlack( required fontColor ){
		return ( arguments.fontColor == 8 ) || ( arguments.fontColor == 32767 );
	}

	private string function fontStyleToCss( required string styleType, required any styleValue, workbook ){
		/*
		Support limited to:
			bold
			color
			italic
			strikethrough
			single underline
		*/
		switch( arguments.styleType ){
			case "bold":
				return "font-weight:" & ( arguments.styleValue? "bold;": "normal;" );
			case "color":
				if( !arguments.KeyExists( "workbook" ) )
					Throw( type=library().getExceptionType() & ".missingRequiredArgument", message="Missing required 'workbook' argument", detail="The 'workbook' argument is required when generating color css styles" );
				//http://ragnarock99.blogspot.co.uk/2012/04/getting-hex-color-from-excel-cell.html
				var rgb = arguments.workbook.getCustomPalette().getColor( arguments.styleValue ).getTriplet();
				var javaColor = CreateObject( "Java", "java.awt.Color" ).init( JavaCast( "int", rgb[ 1 ] ), JavaCast( "int", rgb[ 2 ] ), JavaCast( "int", rgb[ 3 ] ) );
				var hex	=	CreateObject( "Java", "java.lang.Integer" ).toHexString( javaColor.getRGB() );
				hex = hex.subString( 2, hex.length() );
				return "color:##" & hex & ";";
			case "italic":
				return "font-style:" & ( arguments.styleValue? "italic;": "normal;" );
			case "decoration":
				return "text-decoration:#arguments.styleValue#;";//need to pass desired combination of "underline" and "line-through"
		}
		Throw( type=library().getExceptionType() & ".unrecognisedStyle", message="Unrecognised style for css conversion" );
	}

	private numeric function getAWTFontStyle( required any poiFont ){
		var font = library().createJavaObject( "java.awt.Font" );
		var isBold = arguments.poiFont.getBold();
		if( isBold && arguments.poiFont.getItalic() )
			return BitOr( font.BOLD, font.ITALIC );
		if( isBold )
			return font.BOLD;
		if( arguments.poiFont.getItalic() )
			return font.ITALIC;
		return font.PLAIN;
	}

	private numeric function getDefaultCharWidth( required workbook ){
		/*
			Estimates the default character width using Excel's 'Normal' font
			this is a compromise between hard coding a default value and the more complex method of using an AttributedString and TextLayout
		*/
		var defaultFont = arguments.workbook.getFontAt( 0 );
		var style = getAWTFontStyle( defaultFont );
		var font = library().createJavaObject( "java.awt.Font" );
		var javaFont = font.init( defaultFont.getFontName(), style, defaultFont.getFontHeightInPoints() );
		var transform = CreateObject( "java", "java.awt.geom.AffineTransform" );
		var fontContext = CreateObject( "java", "java.awt.font.FontRenderContext" ).init( transform, true, true );
		var bounds = javaFont.getStringBounds( "0", fontContext );
		return bounds.getWidth();
	}

}