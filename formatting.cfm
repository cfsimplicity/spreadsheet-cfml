<cfscript>
private string function richStringCellValueToHtml( required workbook,required cell,required cellValue ){
	var richTextValue=cell.getRichStringCellValue();
	var totalRuns=richTextValue.numFormattingRuns();
	var baseFont=cell.getCellStyle().getFont( workbook );
	if( totalRuns EQ 0  )
		return baseFontToHtml( workbook,cellValue,baseFont );
	// Runs never start at the beginning: the string before the first run is always in the baseFont format
	var startOfFirstRun=richTextValue.getIndexOfFormattingRun( 0 );
	var initialContents=cellValue.Mid( 1,startOfFirstRun );//before the first run
	var initialHtml=baseFontToHtml( workbook,initialContents,baseFont );
	var result=CreateObject( "Java","java.lang.StringBuilder" ).init();
	result.append( initialHtml );
	var endOfCellValuePosition=cellValue.Len();
	for( var runIndex=0; runIndex LT totalRuns; runIndex++ ){
		var run={};
		run.index=runIndex;
		run.number=runIndex+1;
		run.font=workbook.getFontAt( richTextValue.getFontOfFormattingRun( runIndex ) );
		run.css=runFontToHtml( workbook,baseFont,run.font );
		run.isLast = ( run.number EQ totalRuns );
		run.startPosition=richTextValue.getIndexOfFormattingRun( runIndex )+1;
		run.endPosition=run.isLast? endOfCellValuePosition: richTextValue.getIndexOfFormattingRun( runIndex+1 );
		run.length=( ( run.endPosition+1 ) -run.startPosition );
		run.content=cellValue.Mid( run.startPosition,run.length );
		if( run.css.IsEmpty() ){
			result.Append( run.content );
			continue;
		}
		run.html='<span style="#run.css#">#run.content#</span>';
		result.append( run.html );
	}
	return result.toString();
}

private string function runFontToHtml( required workbook,required baseFont,required runFont ){
	/* NB: the order of processing is important for the tests to match */
	var cssStyles=CreateObject( "Java","java.lang.StringBuilder" ).init();
	/* bold */
	if( Compare( runFont.getBold(),baseFont.getBold() ) )
		cssStyles.append( fontStyleToCss( "bold",runFont.getBold() ) );
	/* color */
	if( Compare( runFont.getColor(),baseFont.getColor() ) AND !fontColorIsBlack( runFont.getColor() ) )
		cssStyles.append( fontStyleToCss( "color",runFont.getColor(),workbook ) );
	/* italic */
	if( Compare( runFont.getItalic(),baseFont.getItalic() ) )
		cssStyles.append( fontStyleToCss( "italic",runFont.getItalic() ) );
	/* underline/strike */
	if( Compare( runFont.getStrikeout(),baseFont.getStrikeout() ) OR Compare( runFont.getUnderline(),baseFont.getUnderline() ) ){
		var decorationValue	=	[];
		if( !baseFont.getStrikeout() AND runFont.getStrikeout() )
			decorationValue.Append( "line-through" );
		if( !baseFont.getUnderline() AND runFont.getUnderline() )
			decorationValue.Append( "underline" );
		//if either or both are in the base format, and either or both are NOT in the run format, set the decoration to none.
		if(
				( baseFont.getUnderline() OR baseFont.getStrikeout() )
				AND
				( !runFont.getUnderline() OR !runFont.getUnderline() )
			){
			cssStyles.append( fontStyleToCss( "decoration","none" ) );
		} else {
			cssStyles.append( fontStyleToCss( "decoration",decorationValue.ToList( " " ) ) );
		}
	}
	return cssStyles.toString();
}

private string function baseFontToHtml( required workbook,required contents,required baseFont ){
	/* the order of processing is important for the tests to match */
	/* font family and size not parsed here because all cells would trigger formatting of these attributes: defaults can't be assumed */
	var cssStyles=CreateObject( "Java","java.lang.StringBuilder" ).init();
	/* bold */
	if( baseFont.getBold() )
		cssStyles.append( fontStyleToCss( "bold",true ) );
	/* color */
	if( !fontColorIsBlack( baseFont.getColor() ) )
		cssStyles.append( fontStyleToCss( "color",baseFont.getColor(),workbook ) );
	/* italic */
	if( baseFont.getItalic() )
		cssStyles.append( fontStyleToCss( "italic",true ) );
	/* underline/strike */
	if( baseFont.getStrikeout() OR baseFont.getUnderline() ){
		var decorationValue	=	[];
		if( baseFont.getStrikeout() )
			decorationValue.Append( "line-through" );
		if( baseFont.getUnderline() )
			decorationValue.Append( "underline" );
		cssStyles.append( fontStyleToCss( "decoration",decorationValue.ToList( " " ) ) );
	}
	cssStyles=cssStyles.toString();
	if( cssStyles.IsEmpty() )
		return contents;
	return "<span style=""#cssStyles#"">#contents#</span>";
}

private string function fontStyleToCss( required string styleType,required any styleValue,workbook ){
	/*
	Support limited to:
		bold
		color
		italic
		strikethrough
		underline
	*/
	switch( styleType ){
		case "bold":
			return "font-weight:" & ( styleValue? "bold;": "normal;" );
		case "color":
			if( !arguments.KeyExists( "workbook" ) )
				throw( type=exceptionType,message="The 'workbook' argument is required when generating color css styles" );
			//http://ragnarock99.blogspot.co.uk/2012/04/getting-hex-color-from-excel-cell.html
			var rgb=workbook.getCustomPalette().getColor( styleValue ).getTriplet();
			var javaColor = CreateObject( "Java","java.awt.Color" ).init( JavaCast( "int",rgb[ 1 ] ),JavaCast( "int",rgb[ 2 ] ),JavaCast( "int",rgb[ 3 ] ) );
			var hex	=	CreateObject( "Java","java.lang.Integer" ).toHexString( javaColor.getRGB() );
			hex=hex.subString( 2,hex.length() );
			return "color:##" & hex & ";";
		case "italic":
			return "font-style:" & ( styleValue? "italic;": "normal;" ); 
		case "decoration":
			return "text-decoration:#styleValue#;";//need to pass desired combination of "underline" and "line-through"
	}
	throw( type=exceptionType,message="Unrecognised style for css conversion" );
}

private boolean function fontColorIsBlack( required fontColor ){
	return ( fontColor IS 8 ) OR ( fontColor IS 32767 );
}

private any function buildCellStyle( required workbook,required struct format ){
	/*  TODO: Reuse styles  */
	var cellStyle = workbook.createCellStyle();
	var formatter = workbook.getCreationHelper().createDataFormat();
	var font = 0;
	var setting = 0;
	var settingValue = 0;
	var formatIndex = 0;
	/*
		Valid values of the format struct are:
		* alignment
		* bold
		* bottomborder
		* bottombordercolor
		* color
		* dataformat
		* fgcolor
		* fillpattern
		* font
		* fontsize
		* hidden
		* indent
		* italic
		* leftborder
		* leftbordercolor
		* locked
		* rightborder
		* rightbordercolor
		* rotation
		* strikeout
		* textwrap
		* topborder
		* topbordercolor
		* underline
		* verticalalignment  (added in CF9.0.1)
	 */
	for( var setting in format ){
		settingValue = UCase( format[ setting ] );
		switch( setting ){
			case "alignment":
				cellStyle.setAlignment( cellStyle[ "ALIGN_" & settingValue ] );
			break;
			case "bold":
				font = cloneFont( workbook,workbook.getFontAt( cellStyle.getFontIndex() ) );
				if( format.KeyExists( setting ) )
					font.setBoldweight( font.BOLDWEIGHT_BOLD );
				else
					font.setBoldweight( font.BOLDWEIGHT_NORMAL )
				cellStyle.setFont( font );
			break;
			case "bottomborder":
				cellStyle.setBorderBottom( Evaluate( "cellStyle." & "BORDER_" & UCase( StructFind( format,setting ) ) ) );
			break;
			case "bottombordercolor":
				cellStyle.setBottomBorderColor( JavaCast( "int",getColorIndex( StructFind( format,setting ) ) ) );
			break;
			case "color":
				font = cloneFont( workbook,workbook.getFontAt( cellStyle.getFontIndex() ) );
				font.setColor( getColorIndex( StructFind( format,setting ) ) );
				cellStyle.setFont( font );
			break;
			/*  TODO: this is returning the correct data format index from HSSFDataFormat but doesn't seem to have any effect on the cell. Could be that I'm testing with OpenOffice so I'll have to check things in MS Excel  */
			case "dataformat":
				cellStyle.setDataFormat( formatter.getFormat( JavaCast( "string",format[ setting ] ) ) );
			break;
			case "fgcolor":
				cellStyle.setFillForegroundColor( getColorIndex( StructFind( format,setting ) ) );
				/*  make sure we always apply a fill pattern or the color will not be visible  */
				if( !arguments.KeyExists( "fillpattern" ) )
					cellStyle.setFillPattern( cellStyle.SOLID_FOREGROUND );
			break;
			/*  TODO: CF 9 docs list "nofill" as opposed to "no_fill"; docs wrong? The rest match POI settings exactly.If it really is nofill instead of no_fill, just change to no_fill before calling setFillPattern  */
			case "fillpattern":
				cellStyle.setFillPattern( Evaluate( "cellStyle." & UCase( StructFind( format,setting ) ) ) );
			break;
			case "font":
				font = cloneFont( workbook,workbook.getFontAt( cellStyle.getFontIndex() ) );
				font.setFontName( JavaCast( "string",StructFind( format,setting ) ) );
				cellStyle.setFont( font );
			break;
			case "fontsize":
				font = cloneFont( workbook,workbook.getFontAt( cellStyle.getFontIndex() ) );
				font.setFontHeightInPoints( JavaCast( "int",StructFind( format,setting ) ) );
				cellStyle.setFont( font );
			break;
			/*  TODO: I may just not understand what's supposed to be happening here, but this doesn't seem to do anything */
			case "hidden":
				cellStyle.setHidden( JavaCast( "boolean",StructFind( format, setting ) ) );
			break;
			/*  TODO: I may just not understand what's supposed to be happening here, but this doesn't seem to do anything */
			case "indent":
				cellStyle.setIndention( JavaCast( "int",StructFind( format, setting ) ) );
			break;
			case "italic":
				font = cloneFont( workbook,workbook.getFontAt( cellStyle.getFontIndex ( ) ) );
				if( StructFind( format,setting ) )
					font.setItalic( JavaCast( "boolean",true ) );
				else
					font.setItalic( JavaCast( "boolean",false ) );
				cellStyle.setFont( font );
			break;
			case "leftborder":
				cellStyle.setBorderLeft( Evaluate("cellStyle." & "BORDER_" & UCase( StructFind( format,setting ) ) ) );
			break;
			case "leftbordercolor":
				cellStyle.setLeftBorderColor( getColorIndex( StructFind( format,setting ) ) );
			break;
			/*  TODO: I may just not understand what's supposed to be happening here, but this doesn't seem to do anything */
			case "locked":
				cellStyle.setLocked( JavaCast( "boolean",StructFind( format,setting ) ) );
			break;
			case "rightborder":
				cellStyle.setBorderRight( Evaluate("cellStyle." & "BORDER_" & UCase( StructFind( format,setting ) ) ) );
			break;
			case "rightbordercolor":
				cellStyle.setRightBorderColor( getColorIndex( StructFind( format,setting ) ) );
			break;
			case "rotation":
				cellStyle.setRotation( JavaCast( "int",StructFind( format,setting ) ) );
			break;
			case "strikeout":
				font = cloneFont( workbook,workbook.getFontAt( cellStyle.getFontIndex() ) );
				if( StructFind( format,setting ) )
					font.setStrikeout( JavaCast( "boolean",true ) );
				else
					font.setStrikeout( JavaCast( "boolean",false ) );
				cellStyle.setFont( font );
			break;
			case "textwrap":
				cellStyle.setWrapText( JavaCast( "boolean",StructFind( format,setting ) ) );
			break;
			case "topborder":
				cellStyle.setBorderTop( Evaluate( "cellStyle." & "BORDER_" & UCase( StructFind( format,setting ) ) ) );
			break;
			case "topbordercolor":
				cellStyle.setTopBorderColor( getColorIndex( StructFind( format,setting ) ) );
			break;
			case "underline":
				font = cloneFont( workbook,workbook.getFontAt( cellStyle.getFontIndex() ) );
				if( StructFind( format,setting ) )
					font.setUnderline( JavaCast( "boolean",true ) );
				else
					font.setUnderline( JavaCast( "boolean",false ) );
				cellStyle.setFont( font );
			break;
			case "verticalalignment":
				cellStyle.setVerticalAlignment( cellStyle[ settingValue ] );
			break;
		}
	}
	return cellStyle;
}

private any function cloneFont( required workbook,required fontToClone ){
	var newFont = workbook.createFont();
	/*  copy the existing cell's font settings to the new font  */
	newFont.setBoldweight( fontToClone.getBoldweight() );
	newFont.setCharSet( fontToClone.getCharSet() );
	newFont.setColor( fontToClone.getColor() );
	newFont.setFontHeight( fontToClone.getFontHeight() );
	newFont.setFontName( fontToClone.getFontName() );
	newFont.setItalic( fontToClone.getItalic() );
	newFont.setStrikeout( fontToClone.getStrikeout() );
	newFont.setTypeOffset( fontToClone.getTypeOffset() );
	newFont.setUnderline( fontToClone.getUnderline() );
	return newFont;
}

private numeric function getColorIndex( required string colorName ){
	try{
		var findColor = colorName.Trim().UCase();
		var IndexedColors = CreateObject( "Java","org.apache.poi.ss.usermodel.IndexedColors" );
		var color	= IndexedColors.valueOf( JavaCast( "string",findColor ) );
		return color.getIndex();
	}
	catch( any exception ){
		throw( type=exceptionType,message="Invalid Color",detail="The color provided (#colorName#) is not valid." );
	}
}
</cfscript>