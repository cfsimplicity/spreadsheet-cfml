component extends="base"{

	any function createAnchor( required factory, required struct comment, required struct cellAddress ){
		var anchor = arguments.factory.createClientAnchor();
		var positionSpecified = arguments.comment.KeyExists( "anchor" );
		if( positionSpecified )
			var positions = arguments.comment.anchor.ListToArray();
		// else no position specified, so use the row/column values to set a default
		var anchorValues = {
			col1: positionSpecified? positions[ 1 ]: arguments.cellAddress.column
			,row1: positionSpecified? positions[ 2 ]: arguments.cellAddress.row
			,col2: positionSpecified? positions[ 3 ]: arguments.cellAddress.column+2
			,row2: positionSpecified? positions[ 4 ]: arguments.cellAddress.row+2
		};
		anchor.setCol1( JavaCast( "int", anchorValues.col1 ) );
		anchor.setRow1( JavaCast( "int", anchorValues.row1 ) );
		anchor.setCol2( JavaCast( "int", anchorValues.col2 ) );
		anchor.setRow2( JavaCast( "int", anchorValues.row2 ) );
		return anchor;
	}

	any function addFontStylesToComment( required struct comment, required workbook, required commentString ){
		if( !commentHasFontStyles( arguments.comment ) )
			return this;
		var font = arguments.workbook.createFont();
		if( arguments.comment.KeyExists( "bold" ) )
			font.setBold( JavaCast( "boolean", arguments.comment.bold ) );
		if( arguments.comment.KeyExists( "color" ) )
			font.setColor( getColorHelper().getColor( arguments.workbook, arguments.comment.color ) );
		if( arguments.comment.KeyExists( "font" ) )
			font.setFontName( JavaCast( "string", arguments.comment.font ) );
		if( arguments.comment.KeyExists( "italic" ) )
			font.setItalic( JavaCast( "string", arguments.comment.italic ) );
		if( arguments.comment.KeyExists( "size" ) )
			font.setFontHeightInPoints( JavaCast( "int", arguments.comment.size ) );
		if( arguments.comment.KeyExists( "strikeout" ) )
			font.setStrikeout( JavaCast( "boolean", arguments.comment.strikeout ) );
		if( arguments.comment.KeyExists( "underline" ) )
			font.setUnderline( JavaCast( "byte", arguments.comment.underline ) );
		arguments.commentString.applyFont( font );
		return this;
	}

	any function addHSSFonlyStyles( required struct comment, required commentObject ){
		//the following 5 properties are not currently supported on XSSFComment: https://github.com/cfsimplicity/spreadsheet-cfml/issues/192
		if( arguments.comment.KeyExists( "fillColor" ) ){
			var javaColorRGB = getColorHelper().getJavaColorRGBFor( arguments.comment.fillColor );
			arguments.commentObject.setFillColor(
				JavaCast( "int", javaColorRGB.red )
				,JavaCast( "int", javaColorRGB.green )
				,JavaCast( "int", javaColorRGB.blue )
			);
		}
		if( arguments.comment.KeyExists( "lineStyle" ) )
		 	arguments.commentObject.setLineStyle( JavaCast( "int", arguments.commentObject[ "LINESTYLE_" & arguments.comment.lineStyle.UCase() ] ) );
		if( arguments.comment.KeyExists( "lineStyleColor" ) ){
			var javaColorRGB = getColorHelper().getJavaColorRGBFor( arguments.comment.lineStyleColor );
			arguments.commentObject.setLineStyleColor(
				JavaCast( "int", javaColorRGB.red )
				,JavaCast( "int", javaColorRGB.green )
				,JavaCast( "int", javaColorRGB.blue )
			);
		}
		/* Horizontal alignment can be left, center, right, justify, or distributed. Note that the constants on the Java class are slightly different in some cases: 'center'=CENTERED 'justify'=JUSTIFIED */
		if( arguments.comment.KeyExists( "horizontalAlignment" ) ){
			if( arguments.comment.horizontalAlignment.UCase() == "CENTER" )
				arguments.comment.horizontalAlignment = "CENTERED";
			if( arguments.comment.horizontalAlignment.UCase() == "JUSTIFY" )
				arguments.comment.horizontalAlignment = "JUSTIFIED";
			arguments.commentObject.setHorizontalAlignment( JavaCast( "int", arguments.commentObject[ "HORIZONTAL_ALIGNMENT_" & arguments.comment.horizontalalignment.UCase() ] ) );
		}
		/* Vertical alignment can be top, center, bottom, justify, and distributed. Note that center and justify are DIFFERENT than the constants for horizontal alignment, which are CENTERED and JUSTIFIED. */
		if( arguments.comment.KeyExists( "verticalAlignment" ) )
			arguments.commentObject.setVerticalAlignment( JavaCast( "int", arguments.commentObject[ "VERTICAL_ALIGNMENT_" & arguments.comment.verticalAlignment.UCase() ] ) );
		return this;
	}

	/* Private */

	private boolean function commentHasFontStyles( required struct comment ){
		return (
			arguments.comment.KeyExists( "bold" )
			|| arguments.comment.KeyExists( "color" )
			|| arguments.comment.KeyExists( "font" )
			|| arguments.comment.KeyExists( "italic" )
			|| arguments.comment.KeyExists( "size" )
			|| arguments.comment.KeyExists( "strikeout" )
			|| arguments.comment.KeyExists( "underline" )
		);
	}

}