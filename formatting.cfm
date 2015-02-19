<cfscript>
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
			/*  TODO: this is returning the correct data format index from HSSFDataFormat but 
							doesn't seem to have any effect on the cell. Could be that I'm testing 
							with OpenOffice so I'll have to check things in MS Excel  */
			case "dataformat":
				cellStyle.setDataFormat( formatter.getFormat( JavaCast( "string",format[ setting ] ) ) );
			break;
			case "fgcolor":
				cellStyle.setFillForegroundColor( getColorIndex( StructFind( format,setting ) ) );
				/*  make sure we always apply a fill pattern or the color will not be visible  */
				if( !arguments.KeyExists( "fillpattern" ) )
					cellStyle.setFillPattern( cellStyle.SOLID_FOREGROUND );
			break;
			/*  TODO: CF 9 docs list "nofill" as opposed to "no_fill"; docs wrong? The rest match POI 
							settings exactly.If it really is nofill instead of no_fill, just change to no_fill 
							before calling setFillPattern  */
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