component extends="base"{

	property name="dataFormatter";
	property name="cachedCellStyles" type="struct";//avoid duplication of cellStyles by caching and re-using them for the life of the library

	format function init( required Spreadsheet libraryInstance ){
		initCellStyleCache();
		return super.init( arguments.libraryInstance );
	}

	struct function getCachedCellStyles(){
		return variables.cachedCellStyles;
	}

	format function initCellStyleCache(){
		variables.cachedCellStyles = { xls: {}, xlsx: {} };
		return this;
	}

	any function getDataFormatter(){
		if( IsNull( variables.dataFormatter ) )
			variables.dataFormatter = library().createJavaObject( "org.apache.poi.ss.usermodel.DataFormatter" ).init();
		return variables.dataFormatter;
	}

	any function setCellStyle( required cell, required cellStyle, struct format ){
		try{
			arguments.cell.setCellStyle( arguments.cellStyle );
		}
		catch( any exception ){
			if( !exception.message CONTAINS "Style does not belong to the supplied Workbook" )
				rethrow;
			var workbook = arguments.cell.getSheet().getWorkbook();
			var newCellStyleForThisWorkbook = workbook.createCellStyle();
			newCellStyleForThisWorkbook.cloneStyleFrom( arguments.cellStyle );
			arguments.cell.setCellStyle( newCellStyleForThisWorkbook );
			if( !arguments.KeyExists( "format" ) || StructIsEmpty( arguments.format ) )
				return;
			var spreadsheetType = library().isXmlFormat( workbook )? "xlsx": "xls";
			var cellStyleID = getCellStyleIDfromFormat( arguments.format );
			cacheCellStyle( cellStyleID, spreadsheetType, newCellStyleForThisWorkbook );
		}
	}

	any function getCachedCellStyle( required workbook, required struct format ){
		var spreadsheetType = library().isXmlFormat( arguments.workbook )? "xlsx": "xls";
		var cellStyleID = getCellStyleIDfromFormat( arguments.format );
		if( !variables.cachedCellStyles[ spreadsheetType ].KeyExists( cellStyleID ) ){
			var cellStyle = buildCellStyle( workbook, format );
			cacheCellStyle( cellStyleID, spreadsheetType, cellStyle );
		}
		return variables.cachedCellStyles[ spreadsheetType ][ cellStyleID ];
	}

	any function buildCellStyle( required workbook, required struct format, existingStyle ){
		var cellStyle = arguments.workbook.createCellStyle();
		if( arguments.KeyExists( "existingStyle" ) )
			cellStyle.cloneStyleFrom( arguments.existingStyle );
		for( var setting in arguments.format )
			setCellStyleFromFormatSetting( arguments.workbook, cellStyle, arguments.format, setting );
		return cellStyle;
	}

	boolean function isValidCellStyleObject( required workbook, required any object ){
		if( library().isBinaryFormat( arguments.workbook ) )
			return ( arguments.object.getClass().getCanonicalName() == "org.apache.poi.hssf.usermodel.HSSFCellStyle" );
		return ( arguments.object.getClass().getCanonicalName() == "org.apache.poi.xssf.usermodel.XSSFCellStyle" );
	}

	struct function checkFormatArguments( required workbook, boolean overwriteCurrentStyle=true ){
		if( !arguments.KeyExists( "format" ) && !arguments.KeyExists( "cellStyle" ) )
			Throw( type=library().getExceptionType() & ".missingRequiredArgument", message="Missing argument: 'format'", detail="The 'format' argument is required" );
		if( arguments.KeyExists( "format" ) && IsStruct( arguments.format ) )
			return arguments;
		//assume a cellStyle object has been supplied either as cellStyle or format
		if( !arguments.overwriteCurrentStyle )
			Throw( type=library().getExceptionType() & ".invalidArgumentCombination", message="Invalid argument combination", detail="If you supply a 'cellStyle' the 'overwriteCurrentStyle' cannot be false" );
		if( !arguments.KeyExists( "cellStyle" ) )
			arguments.cellStyle = arguments.format;
		if( !isValidCellStyleObject( arguments.workbook, arguments.cellStyle ) )
			Throw( type=library().getExceptionType() & ".invalidCellStyleArgument", message="Invalid argument", detail="The 'cellStyle' supplied is not a valid POI cellStyle object" );
		return arguments;
	}

	string function patternNameFromIndex( required numeric index ){
		switch( arguments.index ){
			case 0: return "NO_FILL";
			case 1: return "SOLID_FOREGROUND";
			case 2: return "FINE_DOTS";
			case 3: return "ALT_BARS";
			case 4: return "SPARSE_DOTS";
			case 5: return "THICK_HORZ_BANDS";
			case 6: return "THICK_VERT_BANDS";
			case 7: return "THICK_BACKWARD_DIAG";
			case 8: return "THICK_FORWARD_DIAG";
			case 9: return "BIG_SPOTS";
			case 10: return "BRICKS";
			case 11: return "THIN_HORZ_BANDS";
			case 12: return "THIN_VERT_BANDS";
			case 13: return "THIN_BACKWARD_DIAG";
			case 14: return "THIN_FORWARD_DIAG";
			case 15: return "SQUARES";
			case 16: return "DIAMONDS";
			case 17: return "LESS_DOTS";
			case 18: return "LEAST_DOTS";
			default: return "unknown";
		}
	}

	string function underlineNameFromIndex( required numeric index ){
		switch( arguments.index ){
			case 0: return "none";
			case 1: return "single";
			case 2: return "double";
			case 33: return "single accounting";
			case 34: return "double accounting";
			default: return "unknown";
		}
	}

	numeric function underlineIndexFromValue( required any value ){
		switch( arguments.value ){
			case "none": return 0;
			case "single": return 1;
			case "double": return 2;
			case "single accounting": return 33;
			case "double accounting": return 34;
		}
		if( IsBoolean( arguments.value ) )
			return arguments.value? 1: 0;
		return -1;
	}

	string function richStringCellValueToHtml( required workbook, required cell, required cellValue ){
		var richTextValue = arguments.cell.getRichStringCellValue();
		var totalRuns = richTextValue.numFormattingRuns();
		var baseFont = arguments.cell.getCellStyle().getFont( arguments.workbook );
		if( totalRuns == 0  )
			return getFontHelper().baseFontToHtml( arguments.workbook, arguments.cellValue, baseFont );
		// Runs never start at the beginning: the string before the first run is always in the baseFont format
		var startOfFirstRun = richTextValue.getIndexOfFormattingRun( 0 );
		var initialContents = arguments.cellValue.Mid( 1, startOfFirstRun );//before the first run
		var initialHtml = getFontHelper().baseFontToHtml( arguments.workbook, initialContents, baseFont );
		var result = getStringHelper().newJavaStringBuilder();
		result.Append( initialHtml );
		var endOfCellValuePosition = arguments.cellValue.Len();
		for( var runIndex = 0; runIndex < totalRuns; runIndex++ ){
			var run = {};
			run.index = runIndex;
			run.number = ( runIndex +1 );
			run.font = arguments.workbook.getFontAt( richTextValue.getFontOfFormattingRun( runIndex ) );
			run.css = getFontHelper().runFontToHtml( arguments.workbook, baseFont, run.font );
			run.isLast = ( run.number == totalRuns );
			run.startPosition = ( richTextValue.getIndexOfFormattingRun( runIndex ) +1 );
			run.endPosition = run.isLast? endOfCellValuePosition: richTextValue.getIndexOfFormattingRun( ( runIndex +1 ) );
			run.length = ( ( run.endPosition +1 ) -run.startPosition );
			run.content = arguments.cellValue.Mid( run.startPosition, run.length );
			if( run.css.IsEmpty() ){
				result.Append( run.content );
				continue;
			}
			run.html = '<span style="#run.css#">#run.content#</span>';
			result.Append( run.html );
		}
		return result.toString();
	}

	/* Private */

	private string function getCellStyleIDfromFormat( required struct format ){
		return Hash( arguments.format.toString() );
	}

	private void function cacheCellStyle( required string ID, required string spreadsheetType, required cellStyle ){
		variables.cachedCellStyles[ arguments.spreadsheetType ][ arguments.ID ] = arguments.cellStyle;
	}

	private any function setCellStyleFromFormatSetting(
		required workbook
		,required cellStyle
		,required struct format
		,required string setting
	){
		var font = 0;
		var settingValue = arguments.format[ arguments.setting ];
		switch( arguments.setting ){
			case "alignment":
				var alignment = arguments.cellStyle.getAlignment()[ JavaCast( "string", UCase( settingValue ) ) ];
				arguments.cellStyle.setAlignment( alignment );
			return this;
			case "bold":
				font = getFontHelper().cloneFont( arguments.workbook, arguments.workbook.getFontAt( arguments.cellStyle.getFontIndexAsInt() ) );
				font.setBold( JavaCast( "boolean", settingValue ) );
				arguments.cellStyle.setFont( font );
			return this;
			case "bottomborder":
				var borderStyle = arguments.cellStyle.getBorderBottom()[ JavaCast( "string", UCase( settingValue ) ) ];
				arguments.cellStyle.setBorderBottom( borderStyle );
			return this;
			case "bottombordercolor":
				arguments.cellStyle.setBottomBorderColor( getColorHelper().getColor( arguments.workbook, settingValue ) );
			return this;
			case "color":
				font = getFontHelper().cloneFont( arguments.workbook, arguments.workbook.getFontAt( arguments.cellStyle.getFontIndexAsInt() ) );
				font.setColor( getColorHelper().getColor( arguments.workbook, settingValue ) );
				arguments.cellStyle.setFont( font );
			return this;
			case "dataformat":
				var dataFormat = arguments.workbook.getCreationHelper().createDataFormat();
				arguments.cellStyle.setDataFormat( dataFormat.getFormat( JavaCast( "string", settingValue ) ) );
			return this;
			case "fgcolor":
				arguments.cellStyle.setFillForegroundColor( getColorHelper().getColor( arguments.workbook, settingValue ) );
				// make sure we always apply a fill pattern or the color will not be visible
				if( !arguments.format.KeyExists( "fillpattern" ) ){
					var fillpattern = arguments.cellStyle.getFillPattern()[ JavaCast( "string", "SOLID_FOREGROUND" ) ];
					arguments.cellStyle.setFillPattern( fillpattern );
				}
			return this;
			case "fillpattern":
			 //ACF docs list "nofill" as opposed to "no_fill"
				if( settingValue == "nofill" )
					settingValue = "NO_FILL";
				var fillpattern = arguments.cellStyle.getFillPattern()[ JavaCast( "string", UCase( settingValue ) ) ];
				arguments.cellStyle.setFillPattern( fillpattern );
			return this;
			case "font":
				font = getFontHelper().cloneFont( arguments.workbook, arguments.workbook.getFontAt( arguments.cellStyle.getFontIndexAsInt() ) );
				font.setFontName( JavaCast( "string", settingValue ) );
				arguments.cellStyle.setFont( font );
			return this;
			case "fontsize":
				font = getFontHelper().cloneFont( arguments.workbook, arguments.workbook.getFontAt( arguments.cellStyle.getFontIndexAsInt() ) );
				font.setFontHeightInPoints( JavaCast( "int", settingValue ) );
				arguments.cellStyle.setFont( font );
			return this;
			//  TODO: Doesn't seem to do anything/
			case "hidden":
				arguments.cellStyle.setHidden( JavaCast( "boolean", settingValue ) );
			return this;
			case "indent":
				// Only seems to work on MS Excel. XLS limit is 15.
				var indentValue = library().isXmlFormat( arguments.workbook )? settingValue: Min( 15, settingValue );
				arguments.cellStyle.setIndention( JavaCast( "int", indentValue ) );
			return this;
			case "italic":
				font = getFontHelper().cloneFont( arguments.workbook, arguments.workbook.getFontAt( arguments.cellStyle.getFontIndexAsInt ( ) ) );
				font.setItalic( JavaCast( "boolean", settingValue ) );
				arguments.cellStyle.setFont( font );
			return this;
			case "leftborder":
				var borderStyle = arguments.cellStyle.getBorderLeft()[ JavaCast( "string", UCase( settingValue ) ) ];
				arguments.cellStyle.setBorderLeft( borderStyle );
			return this;
			case "leftbordercolor":
				arguments.cellStyle.setLeftBorderColor( getColorHelper().getColor( arguments.workbook, settingValue ) );
			return this;
			// TODO: Doesn't seem to do anything
			case "locked":
				arguments.cellStyle.setLocked( JavaCast( "boolean", settingValue ) );
			return this;
			case "quoteprefixed":
				arguments.cellStyle.setQuotePrefixed( JavaCast( "boolean", settingValue ) );
			return this;
			case "rightborder":
				var borderStyle = arguments.cellStyle.getBorderRight()[ JavaCast( "string", UCase( settingValue ) ) ];
				arguments.cellStyle.setBorderRight( borderStyle );
			return this;
			case "rightbordercolor":
				arguments.cellStyle.setRightBorderColor( getColorHelper().getColor( arguments.workbook, settingValue ) );
			return this;
			case "rotation":
				arguments.cellStyle.setRotation( JavaCast( "int", settingValue ) );
			return this;
			case "strikeout":
				font = getFontHelper().cloneFont( arguments.workbook, arguments.workbook.getFontAt( arguments.cellStyle.getFontIndexAsInt() ) );
				font.setStrikeout( JavaCast( "boolean", settingValue ) );
				arguments.cellStyle.setFont( font );
			return this;
			case "textwrap":
				arguments.cellStyle.setWrapText( JavaCast( "boolean", settingValue ) );
			return this;
			case "topborder":
				var borderStyle = arguments.cellStyle.getBorderTop()[ JavaCast( "string", UCase( settingValue ) ) ];
				arguments.cellStyle.setBorderTop( borderStyle );
			return this;
			case "topbordercolor":
				arguments.cellStyle.setTopBorderColor( getColorHelper().getColor( arguments.workbook, settingValue ) );
			return this;
			case "underline":
				var underlineType = underlineIndexFromValue( settingValue );
				if( underlineType == -1 )
					return this;
				font = getFontHelper().cloneFont( arguments.workbook, arguments.workbook.getFontAt( arguments.cellStyle.getFontIndexAsInt() ) );
				font.setUnderline( JavaCast( "byte", underlineType ) );
				arguments.cellStyle.setFont( font );
			return this;
			case "verticalalignment":
				var alignment = arguments.cellStyle.getVerticalAlignment()[ JavaCast( "string", UCase( settingValue ) ) ];
				arguments.cellStyle.setVerticalAlignment( alignment );
		}
		return this;
	}

}