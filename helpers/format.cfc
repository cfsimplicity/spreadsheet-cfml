component extends="base" accessors="true"{

	property name="dataFormatter" getter="false" setter="false";

	public any function getDataFormatter(){
		if( IsNull( variables.dataFormatter ) )
			variables.dataFormatter = getClassHelper().loadClass( "org.apache.poi.ss.usermodel.DataFormatter" ).init();
		return variables.dataFormatter;
	}

	public any function buildCellStyle( required workbook, required struct format, existingStyle ){
		var cellStyle = arguments.workbook.createCellStyle();
		if( arguments.KeyExists( "existingStyle" ) )
			cellStyle.cloneStyleFrom( arguments.existingStyle );
		for( var setting in arguments.format )
			setCellStyleFromFormatSetting( arguments.workbook, cellStyle, arguments.format, setting );
		return cellStyle;
	}

	public boolean function isValidCellStyleObject( required workbook, required any object ){
		if( library().isBinaryFormat( arguments.workbook ) )
			return ( arguments.object.getClass().getCanonicalName() == "org.apache.poi.hssf.usermodel.HSSFCellStyle" );
		return ( arguments.object.getClass().getCanonicalName() == "org.apache.poi.xssf.usermodel.XSSFCellStyle" );
	}

	public void function checkFormatArguments( required workbook, boolean overwriteCurrentStyle=true ){
		if( arguments.KeyExists( "cellStyle" ) && !arguments.overwriteCurrentStyle )
			Throw( type=library().getExceptionType(), message="Invalid arguments", detail="If you supply a 'cellStyle' the 'overwriteCurrentStyle' cannot be false" );
		if( arguments.KeyExists( "cellStyle" ) && !isValidCellStyleObject( arguments.workbook, arguments.cellStyle ) )
			Throw( type=library().getExceptionType(), message="Invalid argument", detail="The 'cellStyle' argument is not a valid POI cellStyle object" );
	}

	public void function addCellStyleToFormatMethodArgsIfStyleOverwriteAllowed(
		required workbook
		,required boolean overwriteCurrentStyle
		,required struct formatMethodArgs
		,required struct format
	){
		if( arguments.overwriteCurrentStyle )
			arguments.formatMethodArgs.cellStyle = arguments.cellStyle?: buildCellStyle( arguments.workbook, arguments.format );
	}

	public string function lookupUnderlineFormatCode( required cellFont ){
		switch( arguments.cellFont.getUnderline() ){
			case 0: return "none";
			case 1: return "single";
			case 2: return "double";
			case 33: return "single accounting";
			case 34: return "double accounting";
			default: return "unknown";
		}
	}

	public string function richStringCellValueToHtml( required workbook, required cell, required cellValue ){
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

	private void function setCellStyleFromFormatSetting(
		required workbook
		,required cellStyle
		,required struct format
		,required string setting
	){
		var font = 0;
		var settingValue = arguments.format[ setting ];
		switch( arguments.setting ){
			case "alignment":
				var alignment = arguments.cellStyle.getAlignment()[ JavaCast( "string", UCase( settingValue ) ) ];
				arguments.cellStyle.setAlignment( alignment );
			return;
			case "bold":
				font = getFontHelper().cloneFont( arguments.workbook, arguments.workbook.getFontAt( arguments.cellStyle.getFontIndexAsInt() ) );
				font.setBold( JavaCast( "boolean", settingValue ) );
				arguments.cellStyle.setFont( font );
			return;
			case "bottomborder":
				var borderStyle = arguments.cellStyle.getBorderBottom()[ JavaCast( "string", UCase( settingValue ) ) ];
				arguments.cellStyle.setBorderBottom( borderStyle );
			return;
			case "bottombordercolor":
				arguments.cellStyle.setBottomBorderColor( getColorHelper().getColor( arguments.workbook, settingValue ) );
			return;
			case "color":
				font = getFontHelper().cloneFont( arguments.workbook, arguments.workbook.getFontAt( arguments.cellStyle.getFontIndexAsInt() ) );
				font.setColor( getColorHelper().getColor( arguments.workbook, settingValue ) );
				arguments.cellStyle.setFont( font );
			return;
			case "dataformat":
				var dataFormat = arguments.workbook.getCreationHelper().createDataFormat();
				arguments.cellStyle.setDataFormat( dataFormat.getFormat( JavaCast( "string", settingValue ) ) );
			return;
			case "fgcolor":
				arguments.cellStyle.setFillForegroundColor( getColorHelper().getColor( arguments.workbook, settingValue ) );
				// make sure we always apply a fill pattern or the color will not be visible
				if( !arguments.format.KeyExists( "fillpattern" ) ){
					var fillpattern = arguments.cellStyle.getFillPattern()[ JavaCast( "string", "SOLID_FOREGROUND" ) ];
					arguments.cellStyle.setFillPattern( fillpattern );
				}
			return;
			case "fillpattern":
			 //ACF docs list "nofill" as opposed to "no_fill"
				if( settingValue == "nofill" )
					settingValue = "NO_FILL";
				var fillpattern = arguments.cellStyle.getFillPattern()[ JavaCast( "string", UCase( settingValue ) ) ];
				arguments.cellStyle.setFillPattern( fillpattern );
			return;
			case "font":
				font = getFontHelper().cloneFont( arguments.workbook, arguments.workbook.getFontAt( arguments.cellStyle.getFontIndexAsInt() ) );
				font.setFontName( JavaCast( "string", settingValue ) );
				arguments.cellStyle.setFont( font );
			return;
			case "fontsize":
				font = getFontHelper().cloneFont( arguments.workbook, arguments.workbook.getFontAt( arguments.cellStyle.getFontIndexAsInt() ) );
				font.setFontHeightInPoints( JavaCast( "int", settingValue ) );
				arguments.cellStyle.setFont( font );
			return;
			//  TODO: Doesn't seem to do anything/
			case "hidden":
				arguments.cellStyle.setHidden( JavaCast( "boolean", settingValue ) );
			return;
			case "indent":
				// Only seems to work on MS Excel. XLS limit is 15.
				var indentValue = library().isXmlFormat( arguments.workbook )? settingValue: Min( 15, settingValue );
				arguments.cellStyle.setIndention( JavaCast( "int", indentValue ) );
			return;
			case "italic":
				font = getFontHelper().cloneFont( arguments.workbook, arguments.workbook.getFontAt( arguments.cellStyle.getFontIndexAsInt ( ) ) );
				font.setItalic( JavaCast( "boolean", settingValue ) );
				arguments.cellStyle.setFont( font );
			return;
			case "leftborder":
				var borderStyle = arguments.cellStyle.getBorderLeft()[ JavaCast( "string", UCase( settingValue ) ) ];
				arguments.cellStyle.setBorderLeft( borderStyle );
			return;
			case "leftbordercolor":
				arguments.cellStyle.setLeftBorderColor( getColorHelper().getColor( arguments.workbook, settingValue ) );
			return;
			// TODO: Doesn't seem to do anything
			case "locked":
				arguments.cellStyle.setLocked( JavaCast( "boolean", settingValue ) );
			return;
			case "quoteprefixed":
				arguments.cellStyle.setQuotePrefixed( JavaCast( "boolean", settingValue ) );
			return;
			case "rightborder":
				var borderStyle = arguments.cellStyle.getBorderRight()[ JavaCast( "string", UCase( settingValue ) ) ];
				arguments.cellStyle.setBorderRight( borderStyle );
			return;
			case "rightbordercolor":
				arguments.cellStyle.setRightBorderColor( getColorHelper().getColor( arguments.workbook, settingValue ) );
			return;
			case "rotation":
				arguments.cellStyle.setRotation( JavaCast( "int", settingValue ) );
			return;
			case "strikeout":
				font = getFontHelper().cloneFont( arguments.workbook, arguments.workbook.getFontAt( arguments.cellStyle.getFontIndexAsInt() ) );
				font.setStrikeout( JavaCast( "boolean", settingValue ) );
				arguments.cellStyle.setFont( font );
			return;
			case "textwrap":
				arguments.cellStyle.setWrapText( JavaCast( "boolean", settingValue ) );
			return;
			case "topborder":
				var borderStyle = arguments.cellStyle.getBorderTop()[ JavaCast( "string", UCase( settingValue ) ) ];
				arguments.cellStyle.setBorderTop( borderStyle );
			return;
			case "topbordercolor":
				arguments.cellStyle.setTopBorderColor( getColorHelper().getColor( arguments.workbook, settingValue ) );
			return;
			case "underline":
				var underlineType = lookupUnderlineFormat( settingValue );
				if( underlineType == -1 )
					return;
				font = getFontHelper().cloneFont( arguments.workbook, arguments.workbook.getFontAt( arguments.cellStyle.getFontIndexAsInt() ) );
				font.setUnderline( JavaCast( "byte", underlineType ) );
				arguments.cellStyle.setFont( font );
			return;
			case "verticalalignment":
				var alignment = arguments.cellStyle.getVerticalAlignment()[ JavaCast( "string", UCase( settingValue ) ) ];
				arguments.cellStyle.setVerticalAlignment( alignment );
		}
	}

	private numeric function lookupUnderlineFormat( required any formatSettingValue ){
		switch( arguments.formatSettingValue ){
			case "none": return 0;
			case "single": return 1;
			case "double": return 2;
			case "single accounting": return 33;
			case "double accounting": return 34;
		}
		if( IsBoolean( arguments.formatSettingValue ) )
			return arguments.formatSettingValue? 1: 0;
		return -1;
	}

}