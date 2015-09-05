<cfscript>
private void function addInfoBinary( required workbook,required struct info ){
	workbook.createInformationProperties(); // creates the following if missing
	var documentSummaryInfo = workbook.getDocumentSummaryInformation();
	var summaryInfo = workbook.getSummaryInformation();
	for( var key in info ){
		var value = JavaCast( "string",info[ key ] );
		switch( key ){
			case "author":
				summaryInfo.setAuthor( value );
				break;
			case "category":
				documentSummaryInfo.setCategory( value );
				break;
			case "lastauthor":
				summaryInfo.setLastAuthor( value );
				break;
			case "comments":
				summaryInfo.setComments( value );
				break;
			case "keywords":
				summaryInfo.setKeywords( value );
				break;
			case "manager":
				documentSummaryInfo.setManager( value );
				break;
			case "company":
				documentSummaryInfo.setCompany( value );
				break;
			case "subject":
				summaryInfo.setSubject( value );
				break;
			case "title":
				summaryInfo.setTitle( value );
				break;
		}
	}
}

private void function addInfoXml( required workbook,required struct info ){
	var documentProperties = workbook.getProperties().getExtendedProperties().getUnderlyingProperties();
	var coreProperties = workbook.getProperties().getCoreProperties();
	for( var key in info ){
		var value = JavaCast( "string",info[ key ] );
		switch( key ){
			case "author":
				coreProperties.setCreator( value  );
				break;
			case "category":
				coreProperties.setCategory( value );
				break;
			case "lastauthor":
				coreProperties.getUnderlyingProperties().setLastModifiedByProperty( value );
				break;
			case "comments":
				coreProperties.setDescription( value );
				break;
			case "keywords":
				coreProperties.setKeywords( value );
				break;
			case "subject":
				coreProperties.setSubjectProperty( value );
				break;
			case "title":
				coreProperties.setTitle( value );
				break;
			case "manager":
				documentProperties.setManager( value );
				break;
			case "company":
				documentProperties.setCompany( value );
				break;
		}
	}
}

private void function addRowToSheetData( required workbook,required struct sheet,required numeric rowIndex,boolean includeRichTextFormatting=false ){
	if( ( rowIndex EQ sheet.headerRowIndex ) AND !sheet.includeHeaderRow )
		return;
	var rowData=[];
	var row=sheet.object.GetRow( JavaCast( "int",rowIndex ) );
	if( IsNull( row ) ){
		if( sheet.includeBlankRows )
			sheet.data.Append( rowData );
		return;
	}
	if( rowIsEmpty( row ) AND !sheet.includeBlankRows )
		return;
	rowData=getRowData( workbook,row,sheet.columnRanges,includeRichTextFormatting );
	sheet.data.Append( rowData );
	if( !sheet.columnRanges.Len() ){
		var rowColumnCount = row.GetLastCellNum();
		sheet.totalColumnCount = Max( sheet.totalColumnCount,rowColumnCount );
	}
}

private struct function binaryInfo( required workbook ){
	var documentProperties = workbook.getDocumentSummaryInformation();
	var coreProperties = workbook.getSummaryInformation();
	return {
		author = coreProperties.getAuthor()?:""
		,category = documentProperties.getCategory()?:""
		,comments = coreProperties.getComments()?:""
		,creationDate = coreProperties.getCreateDateTime()?:""
		,lastEdited = ( coreProperties.getEditTime() EQ 0 )? "": CreateObject( "java","java.util.Date" ).init( coreProperties.getEditTime() )
		,subject = coreProperties.getSubject()?:""
		,title = coreProperties.getTitle()?:""
		,lastAuthor = coreProperties.getLastAuthor()?:""
		,keywords = coreProperties.getKeywords()?:""
		,lastSaved = coreProperties.getLastSaveDateTime()?:""
		,manager = documentProperties.getManager()?:""
		,company = documentProperties.getCompany()?:""
	};
}

private boolean function cellExists( required workbook,required numeric rowNumber,required numeric columnNumber ){
	var rowIndex = rowNumber-1;
	var columnIndex = columnNumber-1;
	var checkRow = this.getActiveSheet( workbook ).getRow( JavaCast( "int",rowIndex ) );
	return !IsNull( checkRow ) AND !IsNull( checkRow.getCell( JavaCast( "int",columnIndex ) ) );
}

private numeric function columnCountFromRanges( required array ranges ){
	var result=0;
	for( var thisRange in ranges ){
		for( var i=thisRange.startAt; i LTE thisRange.endAt; i++ ){
			result++;
		}
	}
	return result;
}

private function createCell( required row,numeric cellNum=arguments.row.getLastCellNum(),overwrite=true ){
	/* get existing cell (if any)  */
	var cell = row.getCell( JavaCast( "int",cellNum ) );
	if( overwrite AND !IsNull( cell ) )
		arguments.row.removeCell( cell );/* forcibly remove the existing cell  */
	if( overwrite OR IsNull( cell ) )
		cell = row.createCell( JavaCast( "int",cellNum ) );/* create a brand new cell  */
	return cell;
}

private function createRow( required workbook,numeric rowNum=getNextEmptyRow( workbook ),boolean overwrite=true ){
	/* get existing row (if any)  */
	var row = getActiveSheet( workbook ).getRow( JavaCast( "int",rowNum ) );
	if( overwrite AND !IsNull( row ) )
		getActiveSheet( workbook ).removeRow( row ) /* forcibly remove existing row and all cells  */
	if( overwrite OR IsNull( getActiveSheet( workbook ).getRow( JavaCast( "int",rowNum ) ) ) )
		row = getActiveSheet( workbook ).createRow( JavaCast("int", rowNum ) );
	return row;
}

private function createWorkBook( required string sheetName,boolean useXmlFormat=false ){
	this.validateSheetName( sheetName );
	var className = useXmlFormat? "org.apache.poi.xssf.usermodel.XSSFWorkbook": "org.apache.poi.hssf.usermodel.HSSFWorkbook";
	return loadPoi( className ).init();
}

private void function deleteHiddenColumnsFromQuery( required sheet,required query result ){
	var startIndex=( sheet.totalColumnCount-1 );
	for( var colIndex=startIndex; colIndex GTE 0; colIndex-- ){
		if( !sheet.object.isColumnHidden( JavaCast( "integer",colIndex ) ) )
			continue;
		var columnNumber=colIndex+1;
		QueryDeleteColumn( result,sheet.columnNames[ columnNumber ] );
		sheet.totalColumnCount--;
		sheet.columnNames.deleteAt( columnNumber );
	}
}

private void function deleteSheetAtIndex( required workbook,required numeric sheetIndex ){
	workbook.removeSheetAt( JavaCast( "int",sheetIndex ) );
}

private numeric function estimateColumnWidth( required workbook,required any value ){
	/* Estimates approximate column width based on cell value and default character width. */
	/*
	"Excel bases its measurement of column widths on the number of digits (specifically, the number of zeros) in the column, using the Normal style font."
	This function approximates the column width using the number of characters and the default character width in the normal font. POI expresses the width in 1/256 of Excel's character unit. The maximum size in POI is: (255 * 256)
	*/
	var defaultWidth = getDefaultCharWidth( workbook );
	var numOfChars = Len( arguments.value );
	var width = ( numOfChars*defaultWidth+5 ) / ( defaultWidth*256 );
    // Do not allow the size to exceed POI's maximum
	return Min( width,( 255*256 ) );
}

private array function extractRanges( required string rangeList ){
	/*
	A range is a comma-delimited list of ranges, and each value can be either a single number or a range of numbers with a hyphen. Ignores any white space.
	Parses and validates a list of row/column numbers. Returns an array of structures with the keys: startAt, endAt
	*/
	var result = [];
	var rangeTest = "^[0-9]{1,}(-[0-9]{1,})?$";
	var ranges = ListToArray( rangeList );
	for( var thisRange in ranges ){
		/* remove all white space */
		thisRange.REReplace( "\s+","","ALL" );
		if( !REFind( rangeTest,thisRange ) )
			throw( type=exceptionType,message="Invalid range value",detail="The range value '#thisRange#' is not valid." );
		var parts = ListToArray( thisRange,"-" );
		//if this is a single number, the start/endAt values are the same
		var range = {
			startAt = parts[ 1 ]
			,endAt = parts[ parts.Len() ]
		};
		result.Append( range );
	}
	return result;
}

private string function filenameSafe( required string input ){
	var charsToRemove	=	"\|\\\*\/\:""<>~&";
	var result = input.REReplace( "[#charsToRemove#]+","","ALL" ).Left( 255 );
	if( result.isEmpty() )
		return	"renamed"; // in case all chars have been replaced (unlikely but possible)
	return result;
}

private void function fillMergedCellsWithVisibleValue( required workbook,required sheet ){
	if( !sheetHasMergedRegions( sheet ) )
		return;
	for( var regionIndex=0; regionIndex LT sheet.getNumMergedRegions(); regionIndex++ ){
		var region=sheet.getMergedRegion( regionIndex );
		var regionStartRowNumber=region.getFirstRow()+1;
		var regionEndRowNumber=region.getLastRow()+1;
		var regionStartColumnNumber=region.getFirstColumn()+1;
		var regionEndColumnNumber=region.getLastColumn()+1;
		var visibleValue=getCellValue( workbook,regionStartRowNumber,regionStartColumnNumber );
		setCellRangeValue( workbook,visibleValue,regionStartRowNumber,regionEndRowNumber,regionStartColumnNumber,regionEndColumnNumber );
	}
}

private string function generateUniqueSheetName( required workbook ){
	/* Generates a unique sheet name (Sheet1, Sheet2, etecetera). */
	var startNumber = workbook.getNumberOfSheets()+1;
	var maxRetry = startNumber+250;
	for( var sheetNumber=startNumber; sheetNumber LTE maxRetry; sheetNumber++ ){
		var proposedName = "Sheet" & sheetNumber;
		if( !sheetExists( workbook,proposedName ) )
			return proposedName;
	}
	/* this should never happen. but if for some reason it did, warn the action failed and abort */
	throw( type=exceptionType,message="Unable to generate name",detail="Unable to generate a unique sheet name" );
}

private function getActiveSheet( required workbook ){
	return workbook.getSheetAt( JavaCast( "int",workbook.getActiveSheetIndex() ) );
}

private function getActiveSheetName( required workbook ){
	return this.getActiveSheet( workbook ).getSheetName();
}

private numeric function getAWTFontStyle( required any poiFont ){
	var font = loadPOI( "java.awt.Font" );
	var isBold = poiFont.getBoldweight() == poiFont.BOLDWEIGHT_BOLD;
	if( isBold && arguments.poiFont.getItalic() )
  	return BitOr( font.BOLD,font.ITALIC );
	if( isBold )
		return font.BOLD;
	if( poiFont.getItalic() )
		return font.ITALIC;
	return font.PLAIN;
}

private function getCellAt( required workbook,required numeric rowNumber,required numeric columnNumber ){
	if( !cellExists( argumentCollection=arguments ) )
		throw( type=exceptionType,message="Invalid cell",detail="The requested cell [#rowNumber#,#columnNumber#] does not exist in the active sheet" );
	var rowIndex = rowNumber-1;
	var columnIndex = columnNumber-1;
	return getActiveSheet( workbook ).getRow( JavaCast( "int",rowIndex ) ).getCell( JavaCast( "int",columnIndex ) );
}

private function getCellUtil(){
	if( IsNull( variables.cellUtil ) )
		variables.cellUtil = loadPoi( "org.apache.poi.ss.util.CellUtil" );
	return variables.cellUtil;
}

private function getCellValueAsType( required workbook,required cell ){
	/* When getting the value of a cell, it is important to know what type of cell value we are dealing with. If you try to grab the wrong value type, an error might be thrown. For that reason, we must check to see what type of cell we are working with. These are the cell types and they are constants of the cell object itself:

		0 - CELL_TYPE_NUMERIC
		1 - CELL_TYPE_STRING
		2 - CELL_TYPE_FORMULA
		3 - CELL_TYPE_BLANK
		4 - CELL_TYPE_BOOLEAN
		5 - CELL_TYPE_ERROR */

	var cellType = cell.GetCellType();
	/* Get the value of the cell based on the data type. The thing to worry about here is cell forumlas and cell dates. Formulas can be strange and dates are stored as numeric types. Here I will just grab dates as floats and formulas I will try to grab as numeric values. */
	if( cellType EQ cell.CELL_TYPE_NUMERIC ){
		/* Get numeric cell data. This could be a standard number, could also be a date value. */
		var dateUtil = this.getDateUtil();
		if( dateUtil.isCellDateFormatted( cell ) )
			return cell.getDateCellValue();
		return cell.getNumericCellValue();
	}
	if( cellType EQ cell.CELL_TYPE_FORMULA ){
		var formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
		return this.getFormatter().formatCellValue( cell,formulaEvaluator );
	}
	if( cellType EQ cell.CELL_TYPE_BOOLEAN )
		return cell.getBooleanCellValue();
 	if( cellType EQ cell.CELL_TYPE_BLANK )
		return "";
	try{
		return cell.getStringCellValue();
	}
	catch( any exception ){
		return "";
	}
}

private function getDateUtil(){
	if( IsNull( variables.dateUtil ) )
		variables.dateUtil = loadPoi( "org.apache.poi.ss.usermodel.DateUtil" );
	return variables.dateUtil;
}

private string function getDateTimeValueFormat( required any value ){
	/* Returns the default date mask for the given value: DATE (only), TIME (only) or TIMESTAMP */
	var dateTime = ParseDateTime( value );
	var dateOnly = CreateDate( Year( dateTime ),Month( dateTime ),Day( dateTime ) );
	if( DateCompare( value,dateOnly,"s" ) EQ 0 )
		return variables.dateFormats.DATE;
	if( DateCompare( "1899-12-30",dateOnly,"d" ) EQ 0 )
		return variables.dateFormats.TIME;
	return variables.dateFormats.TIMESTAMP;
}

private numeric function getDefaultCharWidth( required workbook ){
	/* Estimates the default character width using Excel's 'Normal' font */
	/* this is a compromise between hard coding a default value and the more complex method of using an AttributedString and TextLayout */
	var defaultFont = workbook.getFontAt( 0 );
	var style = getAWTFontStyle( defaultFont );
	var font = loadPOI( "java.awt.Font" );
	var javaFont = font.init( defaultFont.getFontName(),style,defaultFont.getFontHeightInPoints() );
	// this works
	var transform = CreateObject( "java","java.awt.geom.AffineTransform" );
	var fontContext = CreateObject( "java","java.awt.font.FontRenderContext" ).init( transform,true,true );
	var bounds = javaFont.getStringBounds( "0",fontContext );
	return bounds.getWidth();
}

private numeric function getFirstRowNum( required workbook ){
	var firstRow = getActiveSheet( workbook ).getFirstRowNum();
	if( firstRow EQ 0 AND getActiveSheet( workbook ).getPhysicalNumberOfRows() EQ 0 )
		return -1;
	return firstRow;
}

private function getFormatter(){
	/* Returns cell formatting utility object ie org.apache.poi.ss.usermodel.DataFormatter */
	if( IsNull( variables.dataFormatter ) )
		variables.dataFormatter = this.loadPOI( "org.apache.poi.ss.usermodel.DataFormatter" ).init();
	return dataFormatter;
}

private struct function getJavaColorRGB( required string colorName ){
	/* Returns a struct containing RGB values from java.awt.Color for the color name passed in */
	var findColor = colorName.Trim().UCase();
	var color = CreateObject( "Java","java.awt.Color" );
	if( IsNull( color[ findColor ] ) OR !IsInstanceOf( color[ findColor ],"java.awt.Color" ) )//don't use member functions on color
		throw( type=exceptionType,message="Invalid color",detail="The color provided (#colorName#) is not valid." );
	color = color[ findColor ];
	var colorRGB = {
		red = color.getRed()
		,green = color.getGreen()
		,blue = color.getBlue()
	};
	return colorRGB;
}

private numeric function getLastRowNum( required workbook ){
	var lastRow = getActiveSheet( workbook ).getLastRowNum();
	if( lastRow EQ 0 AND getActiveSheet( workbook ).getPhysicalNumberOfRows() EQ 0 )
		return -1;//The sheet is empty. Return -1 instead of 0
	return lastRow;
}

private numeric function getNextEmptyRow( workbook ){
	return getLastRowNum( workbook )+1;
}

private array function getQueryColumnFormats( required workbook,required query query ){
	/* extract the query columns and data types  */
	//var cell	  	= CreateObject( "Java","org.apache.poi.ss.usermodel.Cell" );
	var formatter	= workbook.getCreationHelper().createDataFormat();
	var metadata 	= GetMetaData( query );
	/* assign default formats based on the data type of each column */
	for( var col in metadata ){
		switch( col.typeName ){
			/* apply basic formatting to dates and times for increased readability */
			case "DATE": case "TIMESTAMP":
				col.cellDataType = "DATE";
				col.defaultCellStyle 	= this.buildCellStyle( workbook,{ dataFormat = variables.dateFormats[ col.typeName ] } );
			break;
			case "TIME":
				col.cellDataType = "TIME";
				col.defaultCellStyle 	= this.buildCellStyle( workbook,{ dataFormat = variables.dateFormats[ col.typeName ] } );
			break;
			/* Note: Excel only supports "double" for numbers. Casting very large DECIMIAL/NUMERIC or BIGINT values to double may result in a loss of precision or conversion to NEGATIVE_INFINITY / POSITIVE_INFINITY. */
			case "DECIMAL": case "BIGINT": case "NUMERIC": case "DOUBLE": case "FLOAT": case "INTEGER": case "REAL": case "SMALLINT": case "TINYINT":
				col.cellDataType = "DOUBLE";
			break;
			case "BOOLEAN": case "BIT":
				col.cellDataType = "BOOLEAN";
			break;
			default:
				col.cellDataType = "STRING";
		}
	}
	return metadata;
}

private array function getRowData( required workbook,required row,array columnRanges=[],boolean includeRichTextFormatting=false ){
	var result=[];
	if( !columnRanges.Len() ){
		var columnRange={
			startAt=1
			,endAt=row.GetLastCellNum()
		};
		arguments.columnRanges=[ columnRange ];
	}
	for( var thisRange in columnRanges ){
		for( var i=thisRange.startAt; i LTE thisRange.endAt; i++ ){
			var colIndex=i-1;
			var cell = row.GetCell( JavaCast( "int",colIndex ) );
			if( IsNull( cell ) ){
				result.Append( "" );
				continue;
			}
			var cellValue = this.getCellValueAsType( workbook,cell );
			if( includeRichTextFormatting AND ( cell.GetCellType() EQ cell.CELL_TYPE_STRING ) )
				cellValue = richStringCellValueToHtml( workbook,cell,cellValue );
			result.Append( cellValue );
		}
	}
	return result;
}

private numeric function getSheetIndexFromName( required workbook,required string sheetName ){
	//returns -1 if non-existent
	return workbook.getSheetIndex( JavaCast( "string",sheetName ) );
}

private function initializeCell( required workbook,required numeric rowNumber,required numeric columnNumber ){
	var rowIndex = JavaCast( "int",rowNumber-1 );
	var columnIndex = JavaCast( "int",columnNumber-1 );
	var rowObject = getCellUtil().getRow( rowIndex,getActiveSheet( workbook ) );
	var cellObject = getCellUtil().getCell( rowObject,columnIndex );
	return cellObject;
}

private boolean function isDateObject( required input ){
	return input.getClass().getName() IS "java.util.Date";
}

private boolean function isString( required input ){
	return input.getClass().getName() IS "java.lang.String";
}

private function loadPoi( required string javaclass ){
	if( !server.KeyExists( poiLoaderName ) ){
		var paths = [];
		var libPath = ExpandPath( GetDirectoryFromPath( GetCurrentTemplatePath() ) & "lib/" );
		paths.Append( libPath & "poi-3.12-20150511.jar" );
		paths.Append( libPath & "poi-ooxml-3.12-20150511.jar" );
		paths.Append( libPath & "poi-ooxml-schemas-3.12-20150511.jar" );
		paths.Append( libPath & "xmlbeans-2.6.0.jar" );
		if( !server.KeyExists( poiLoaderName ) ){
			server[ poiLoaderName ] = CreateObject( "component","javaLoader.JavaLoader" ).init( loadPaths=paths,loadColdFusionClassPath=true,trustedSource=true );
		}
	}
	return server[ poiLoaderName ].create( arguments.javaclass );
}

private void function moveSheet( required workbook,required string sheetName,required string moveToIndex ){
	workbook.setSheetOrder( JavaCast( "String",sheetName ),JavaCast( "int",moveToIndex ) );
}

private array function parseRowData( required string line,required string delimiter,boolean handleEmbeddedCommas=true ){
	var elements = ListToArray( arguments.line,arguments.delimiter );
	var potentialQuotes = 0;
	arguments.line = ToString( arguments.line );
	if( arguments.delimiter EQ "," AND arguments.handleEmbeddedCommas )
		potentialQuotes = arguments.line.replaceAll("[^']", "").length();
	if (potentialQuotes <= 1)
	  return elements;
	//For ACF compatibility, find any values enclosed in single quotes and treat them as a single element.
	var currentValue = 0;
	var nextValue = "";
	var isEmbeddedValue = false;
	var values = [];
	var buffer = CreateObject( "Java","java.lang.StringBuilder" ).init();
	var maxElements = ArrayLen( elements );

	for( var i=1; i LTE maxElements; i++) {
	  currentValue = Trim( elements[ i ] );
	  nextValue = i < maxElements ? elements[ i + 1 ] : "";
	  var isComplete = false;
	  var hasLeadingQuote = currentValue.hasPrefix( "'" );
	  var hasTrailingQuote = currentValue.hasSuffix( "'" );
	  var isFinalElement = ( i==maxElements );
	  if( hasLeadingQuote )
		  isEmbeddedValue = true;
	  if( isEmbeddedValue AND hasTrailingQuote )
		  isComplete = true;
	  /* We are finished with this value if:
		  * no quotes were found OR
		  * it is the final value OR
		  * the next value is embedded in quotes
	  */
	  if( !isEmbeddedValue || isFinalElement || nextValue.hasPrefix( "'" ) )
		  isComplete = true;
	  if( isEmbeddedValue || isComplete ){
		  // if this a partial value, append the delimiter
		  if( isEmbeddedValue AND buffer.length() GT 0 )
			  buffer.append( "," );
		  buffer.append( elements[i] );
	  }
	  if( isComplete ){
		  var finalValue = buffer.toString();
		  var startAt = finalValue.indexOf( "'" );
		  var endAt = finalValue.lastIndexOf( "'" );
		  if( isEmbeddedValue AND startAt GTE 0 AND endAt GT startAt )
			  finalValue = finalValue.substring( startAt+1,endAt );
		  values.add( finalValue );
		  buffer.setLength( 0 );
		  isEmbeddedValue = false;
	  }
  }
  return values;
}

private string function queryToCsv( required query query,numeric headerRow,boolean includeHeaderRow ){
	var result=CreateObject( "Java","java.lang.StringBuilder" ).init();
	var crlf=Chr( 13 ) & Chr( 10 );
	var columns=query.ColumnArray();
	var hasHeaderRow=( arguments.KeyExists( "headerRow" ) AND Val( headerRow ) );
	if( hasHeaderRow )
		result.Append( generateCsvRow( columns ) );
	for( var row in query ){
		var rowValues=[];
		for( column in columns ){
			rowValues.Append( row[ column ] );
		}
		result.append( crlf & generateCsvRow( rowValues ) );
	}
	return result.toString().Trim();
}

private string function generateCsvRow( required array values,delimiter="," ){
	var result=CreateObject( "Java","java.lang.StringBuilder" ).init();
	for( var value in values ){
		if( this.isDateObject( value ) ){
			value= DateTimeFormat( value,dateFormats.DATETIME );
		}
		value=Replace( value,'"','""',"ALL" );//can't use member function in case its a non-string
		result.append( '#delimiter#"#value#"' );
	}
	return result.toString().substring( 1 );
}

private string function queryToHtml( required query query,numeric headerRow,boolean includeHeaderRow ){
	var result=CreateObject( "Java","java.lang.StringBuilder" ).init();
	var columns=query.ColumnArray();
	var hasHeaderRow=( arguments.KeyExists( "headerRow" ) AND Val( headerRow ) );
	if( hasHeaderRow ){
		result.append( "<thead>" );
		result.append( generateHtmlRow( columns,true ) );
		result.append( "</thead>" );
	}
	result.append( "<tbody>" );
	for( var row in query ){
		var rowValues=[];
		for( column in columns ){
			rowValues.Append( row[ column ] );
		}
		result.append( generateHtmlRow( rowValues ) );
	}
	result.append( "</tbody>" );
	return result.toString();
}

private string function generateHtmlRow( required array values,boolean isHeader=false ){
	var result=CreateObject( "Java","java.lang.StringBuilder" ).init();
	result.append( "<tr>" );
	var columnTag=isHeader? "th": "td";
	for( var value in values ){
		if( this.isDateObject( value ) ){
			value= DateTimeFormat( value,dateFormats.DATETIME );
		}
		result.append( "<#columnTag#>#value#</#columnTag#>" );
	}
	result.append( "</tr>" );
	return result.toString();
}

private boolean function rowIsEmpty( required row ){
	for( var i=row.getFirstCellNum(); i LT row.getLastCellNum(); i++ ){
    var cell = row.getCell( i );
    if( !IsNull( cell ) && ( cell.getCellType() != cell.CELL_TYPE_BLANK ) )
      return false;
  }
  return true;
}

private void function setCellValueAsType( required workbook,required cell,required value ){
	if( IsNumeric( value ) AND !REFind( value,"^0[\d]+" ) ){ /*  skip numeric strings with leading zeroes. treat those as text  */
		/*  NUMERIC  */
		cell.setCellType( cell.CELL_TYPE_NUMERIC );
		cell.setCellValue( JavaCast( "double",Val( value ) ) );
		return;
	}
	if( IsDate( value ) ){
		/*  DATE  */
		var cellFormat = this.getDateTimeValueFormat( value );
		cell.setCellStyle( this.buildCellStyle( workbook,{ dataFormat=cellFormat } ) );
		cell.setCellType( cell.CELL_TYPE_NUMERIC );
		/*  Excel's uses a different epoch than CF (1900-01-01 versus 1899-12-30). "Time" only values will not display properly without special handling - */
		if( cellFormat EQ variables.dateFormats.TIME ){
			var dateUtil = this.getDateUtil();
			value = TimeFormat( value, "HH:MM:SS" );
		 	cell.setCellValue( dateUtil.convertTime( value ) );
		} else {
			cell.setCellValue( ParseDateTime( value ) );
		}
		return;
	}
	if( IsBoolean( value ) ){
		/* BOOLEAN */
		cell.setCellType( cell.CELL_TYPE_BOOLEAN );
		cell.setCellValue( JavaCast( "boolean",value ) );
		return;
	}
	if( !value.Trim().Len() ){
		/* EMPTY */
		cell.setCellType( cell.CELL_TYPE_BLANK );
		cell.setCellValue( "" );
		return;
	}
	/* STRING */
	cell.setCellType( cell.CELL_TYPE_STRING );
	cell.setCellValue( JavaCast( "string",value ) );
}

private boolean function sheetExists( required workbook,string sheetName,numeric sheetNumber ){
	validateSheetNameOrNumberWasProvided( argumentCollection=arguments );
	if( arguments.KeyExists( "sheetName" ) )
		arguments.sheetNumber = this.getSheetIndexFromName( workbook,sheetName )+1;
		//the position is valid if it an integer between 1 and the total number of sheets in the workbook
	if( sheetNumber AND ( sheetNumber EQ Round( sheetNumber ) ) AND ( sheetNumber LTE workbook.getNumberOfSheets() ) )
		return true;
	return false;
}

private boolean function sheetHasMergedRegions( required sheet ){
	return ( sheet.getNumMergedRegions() GT 0 );
}

private query function sheetToQuery(
	required workbook
	,string sheetName
	,numeric sheetNumber=1
	,numeric headerRow
	,boolean includeHeaderRow=false
	,boolean includeBlankRows=false
	,boolean includeHiddenColumns=false
	,boolean fillMergedCellsWithVisibleValue=false
	,boolean includeRichTextFormatting=false
	,string rows //range
	,string columns //range
	,string columnNames
){
	var sheet={
		includeHeaderRow=includeHeaderRow
		,hasHeaderRow=( arguments.KeyExists( "headerRow" ) AND Val( headerRow ) )
		,includeBlankRows=includeBlankRows
		,columnNames=[]
		,columnRanges=[]
		,totalColumnCount=0
	};
	sheet.headerRowIndex = sheet.hasHeaderRow? headerRow-1: -1;
	if( arguments.KeyExists( "columns" ) ){
		sheet.columnRanges=this.extractRanges( arguments.columns );
		sheet.totalColumnCount=columnCountFromRanges( sheet.columnRanges );
	}
	if( arguments.KeyExists( "sheetName" ) ){
		validateSheetExistsWithName( workbook,sheetName );
		arguments.sheetNumber = getSheetIndexFromName( workbook,sheetName )+1;
	}
	sheet.object = workbook.GetSheetAt( JavaCast( "int",sheetNumber-1 ) );
	if( fillMergedCellsWithVisibleValue )
		this.fillMergedCellsWithVisibleValue( workbook,sheet.object );
	sheet.data = [];
	if( arguments.KeyExists( "rows" ) ){
		var allRanges = this.extractRanges( arguments.rows );
		for( var thisRange in allRanges ){
			for( var rowNumber=thisRange.startAt; rowNumber LTE thisRange.endAt; rowNumber++ ){
				var rowIndex=rowNumber-1;
				this.addRowToSheetData( workbook,sheet,rowIndex,includeRichTextFormatting );
			}
		}
	} else {
		var lastRowIndex = sheet.object.GetLastRowNum();// zero based
		for( var rowIndex=0; rowIndex LTE lastRowIndex; rowIndex++ ){
			this.addRowToSheetData( workbook,sheet,rowIndex,includeRichTextFormatting );
		}
	}
	//generate the query columns
	if( arguments.KeyExists( "columnNames" ) AND arguments.columnNames.Len() ){
		arguments.columnNames=arguments.columnNames.ListToArray();
		for( var i=1; i LTE sheet.totalColumnCount; i++ ){
			var columnName=columnNames[ i ]?: "column" & i;
			sheet.columnNames.Append( columnName );
		}
	} else if( sheet.hasHeaderRow ){
		var headerRow=sheet.object.GetRow( JavaCast( "int",sheet.headerRowIndex ) );
		var rowData=getRowData( workbook,headerRow,sheet.columnRanges );
		var i=1;
		for( var value in rowData ){
			var columnName="column" & i;
			if( this.isString( value ) AND value.Len() )
				columnName=value;
			sheet.columnNames.Append( columnName );
			i++;
		}
	} else {
		for( var i=1; i LTE sheet.totalColumnCount; i++ ){
			sheet.columnNames.Append( "column" & i );
		}
	}
	var result=QueryNew( sheet.columnNames,"",sheet.data );
	if( !includeHiddenColumns ){
		deleteHiddenColumnsFromQuery( sheet,result );
		if( sheet.totalColumnCount EQ 0 )
			return Query();// all columns were hidden: return a blank query.
	}
	return result;
}

void function toggleColumnHidden( required workbook,required numeric columnNumber, required boolean state ){
	var sheet=this.getActiveSheet( workbook );
	sheet.setColumnHidden( JavaCast( "integer",columnNumber-1 ),JavaCast( "boolean",state ) );
}

private void function validateSheetExistsWithName( required workbook,required string sheetName ){
	if( !this.sheetExists( workbook=workbook,sheetName=sheetName ) )
		throw( type=exceptionType,message="Invalid sheet name [#sheetName#]",detail="The specified sheet was not found in the current workbook." );
}

private void function validateSheetNumber( required workbook,required numeric sheetNumber ){
	if( !this.sheetExists( workbook=workbook,sheetNumber=sheetNumber ) ){
		var sheetCount = workbook.getNumberOfSheets();
		throw( type=exceptionType,message="Invalid sheet number [#sheetNumber#]",detail="The sheetNumber must a whole number between 1 and the total number of sheets in the workbook [#sheetCount#]" );
	}
}

private void function validateSheetName( required string sheetName ){
	var poiTool = loadPoi( "org.apache.poi.ss.util.WorkbookUtil" );
	try{
		poiTool.validateSheetName( JavaCast( "String",sheetName ) );
	}
	catch( "java.lang.IllegalArgumentException" exception ){
		throw( type=exceptionType,message="Invalid characters in sheet name",detail=exception.message );
	}
}

private void function validateSheetNameOrNumberWasProvided(){
	if( !arguments.KeyExists( "sheetName" ) AND !arguments.KeyExists( "sheetNumber" ) )
		throw( type=exceptionType,message="Missing Required Argument", detail="Either sheetName or sheetNumber must be provided" );
	if( arguments.KeyExists( "sheetName" ) AND arguments.KeyExists( "sheetNumber" ) )
		throw( type=exceptionType,message="Too Many Arguments", detail="Only one argument is allowed. Specify either a sheetName or sheetNumber, not both" );
}

private function workbookFromFile( required string path ){
	// works with both xls and xlsx
	try{
		lock name="#path#" timeout=5{
			var file = CreateObject( "java","java.io.FileInputStream" ).init( path );
			var workbook = loadPoi( "org.apache.poi.ss.usermodel.WorkbookFactory" ).create( file );
		}
		return workbook;
	}
	finally{
		file.close();
	}
}

private struct function xmlInfo( required workbook ){
	var documentProperties = workbook.getProperties().getExtendedProperties().getUnderlyingProperties();
	var coreProperties = workbook.getProperties().getCoreProperties();
	return {
		author = coreProperties.getCreator()?:""
		,category = coreProperties.getCategory()?:""
		,comments = coreProperties.getDescription()?:""
		,creationDate = coreProperties.getCreated()?:""
		,lastEdited = coreProperties.getModified()?:""
		,subject = coreProperties.getSubject()?:""
		,title = coreProperties.getTitle()?:""
		,lastAuthor = coreProperties.getUnderlyingProperties().getLastModifiedByProperty().getValue()?:""
		,keywords = coreProperties.getKeywords()?:""
		,lastSaved = ""// not available in xml
		,manager = documentProperties.getManager()?:""
		,company = documentProperties.getCompany()?:""
	};
}
</cfscript>
