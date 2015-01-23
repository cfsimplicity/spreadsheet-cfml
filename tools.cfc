component access="package"{

	function init( required formatting,required struct defaultFormats,required string exceptionType ){
		variables.formatting = formatting;
		variables.defaultFormats	=	defaultFormats;
		variables.exceptionType	=	exceptionType;
		return this;
	}

	/* Workaround for an issue with autoSizeColumn(). It does not seem to handle 
		date cells properly. It measures the length of the date "number", instead of 
		the  visible date string ie mm//dd/yyyy. As a result columns are too narrow */
	void function autoSizeColumnFix(
		required workbook
		,required numeric columnIndex /* Base-0 */
		,boolean isDateColumn=false
		,string dateMask=variables.defaultFormats[ "TIMESTAMP" ]
	){
		if( isDateColumn ){
			newWidth = estimateColumnWidth( dateMask & "00000" );
			getActiveSheet( workbook ).setColumnWidth( columnIndex,newWidth );
		} else {
			getActiveSheet( workbook ).autoSizeColumn( JavaCast( "int",columnIndex),true );
		}
	}

	function createCell( required row,numeric cellNum=arguments.row.getLastCellNum(),overwrite=true ){
		/* get existing cell (if any)  */
		var cell = row.getCell( JavaCast( "int",cellNum ) );
		if( overwrite AND !IsNull( cell ) )
			arguments.row.removeCell( cell );/* forcibly remove the existing cell  */
		if( overwrite OR IsNull( cell ) )
			cell = row.createCell( JavaCast( "int",cellNum ) );/* create a brand new cell  */
		return cell;
	}

	function createRow( required workbook,numeric rowNum=getNextEmptyRow( workbook ),boolean overwrite=true ){
		/* get existing row (if any)  */
		var row = getActiveSheet( workbook ).getRow( JavaCast( "int",rowNum ) );
		if( overwrite AND !IsNull( row ) )
			getActiveSheet( workbook ).removeRow( row ) /* forcibly remove existing row and all cells  */
		if( overwrite OR IsNull( getActiveSheet( workbook ).getRow( JavaCast( "int",rowNum ) ) ) )
			row = getActiveSheet( workbook ).createRow( JavaCast("int", rowNum ) );
		return row;
	}

	function createSheet( required workbook,required string sheetName ){
		newSheet = workbook.createSheet( JavaCast( "String", sheetName ) );
		return newSheet;
	}

	function createWorkBook( required string sheetName ){
		return CreateObject( "Java","org.apache.poi.hssf.usermodel.HSSFWorkbook" );
	}

	string function filenameSafe( required string input ){
		var charsToRemove	=	"\|\\\*\/\:""<>~&";
		var result = input.REReplace( "[#charsToRemove#]+","","ALL" ).Left( 255 );
		if( result.isEmpty() )
			return	"renamed"; // in case all chars have been replaced (unlikely but possible)
		return result;
	}

	function getActiveSheet( required workbook ){
		return workbook.getSheetAt( JavaCast( "int",workbook.getActiveSheetIndex() ) );
	} 

	function getCellUtil(){
		if( IsNull( variables.cellUtil ) )
			variables.cellUtil = CreateObject( "Java","org.apache.poi.ss.util.CellUtil" );
		return variables.cellUtil;
	}

	function getDateUtil(){
		if( IsNull( variables.dateUtil ) )
			variables.dateUtil = CreateObject( "Java","org.apache.poi.ss.usermodel.DateUtil" );
		return variables.dateUtil;
	}

	string function getDateTimeValueFormat( required any value ){
		/* Returns the default date mask for the given value: DATE (only), TIME (only) or TIMESTAMP */
		var dateTime = ParseDateTime( value );
		var dateOnly = CreateDate( Year( dateTime ),Month( dateTime ),Day( dateTime ) );
		if( DateCompare( value,dateOnly,"s" ) EQ 0 )
			return variables.defaultFormats.DATE;
		if( DateCompare( "1899-12-30",dateOnly,"d" ) EQ 0 )
			return variables.defaultFormats.TIME;
		return variables.defaultFormats.TIMESTAMP;
	}

	numeric function getFirstRowNum( required workbook ){
		var firstRow = getActiveSheet( workbook ).getFirstRowNum();
		if( firstRow EQ 0 AND getActiveSheet( workbook ).getPhysicalNumberOfRows() EQ 0 )
			return -1;
		return firstRow;
	}

	numeric function getLastRowNum( required workbook ){
		var lastRow = getActiveSheet( workbook ).getLastRowNum();
		if( lastRow EQ 0 AND getActiveSheet( workbook ).getPhysicalNumberOfRows() EQ 0 )
			return -1;//The sheet is empty. Return -1 instead of 0
		return lastRow;
	}

	numeric function getNextEmptyRow( workbook ){
		return getLastRowNum( workbook )+1;
	}

	array function getQueryColumnFormats( required workbook,required query query ){
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
					col.defaultCellStyle 	= formatting.buildCellStyle( workbook,{ dataFormat = variables.defaultFormats[ col.typeName ] } );
				break;
				case "TIME":
					col.cellDataType = "TIME";
					col.defaultCellStyle 	= formatting.buildCellStyle( workbook,{ dataFormat = variables.defaultFormats[ col.typeName ] } );
				break;
				/* Note: Excel only supports "double" for numbers. Casting very large DECIMIAL/NUMERIC
				    or BIGINT values to double may result in a loss of precision or conversion to 
					NEGATIVE_INFINITY / POSITIVE_INFINITY. */
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

	function initializeCell( required workbook,required numeric row,required numeric column ){
		var jRow = JavaCast( "int",row-1 );
		var jColumn = JavaCast( "int",column-1 );
		var rowObject = getCellUtil().getRow( jRow,getActiveSheet( workbook ) );
		var cellObject = getCellUtil().getCell( rowObject,jColumn );
		return cellObject; 
	}

	array function parseRowData( required string line,required string delimiter,boolean handleEmbeddedCommas=true ){
		var elements = ListToArray( arguments.line,arguments.delimiter );
		var potentialQuotes = 0;
		arguments.line = ToString( arguments.line );
		if( arguments.delimiter EQ "," AND arguments.handleEmbeddedCommas )
			potentialQuotes = arguments.line.replaceAll("[^']", "").length();		
		if (potentialQuotes <= 1)
		  return elements;
		/*
			For ACF compatibility, find any values enclosed in single 
			quotes and treat them as a single element.
		*/ 
  	var currentValue = 0;
  	var nextValue = "";
		var isEmbeddedValue = false;
		var values = [];
		var buffer = CreateObject( "Java","java.lang.StringBuilder").init();
		var maxElements = ArrayLen( elements );
		
		for( var i=1; i LTE maxElements; i++) {
		  currentValue = Trim( elements[ i ] );
		  nextValue = i < maxElements ? elements[ i + 1 ] : "";
		  var isComplete = false;
		  var hasLeadingQuote = currentValue.startsWith( "'" );
		  var hasTrailingQuote = currentValue.endsWith( "'" );
		  var isFinalElement = ( i==maxElements );
		  if( hasLeadingQuote )
			  isEmbeddedValue = true;
		  if( isEmbeddedValue AND hasTrailingQuote )
			  isComplete = true;
		  // We are finished with this value if:  
		  // * no quotes were found OR
		  // * it is the final value OR
		  // * the next value is embedded in quotes
		  if( !isEmbeddedValue || isFinalElement || nextValue.startsWith( "'" ) )
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
			  buffer.setLength(0);
			  isEmbeddedValue = false;
		  }	  
	  }
	  return values;
	}

	function setActiveSheet( required workbook,string sheetName ){
		var sheetIndex = workbook.getSheetIndex( JavaCast( "string", sheetName ) ) + 1;
		workbook.setActiveSheet( JavaCast( "int",sheetIndex - 1 ) );
	}

}