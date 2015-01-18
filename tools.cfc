component access="package"{

	function init( required workbook,required string exceptionType ){
		variables.workbook	=	workbook;
		variables.exceptionType	=	exceptionType;
		return this;
	}

	/* Workaround for an issue with autoSizeColumn(). It does not seem to handle 
		date cells properly. It measures the length of the date "number", instead of 
		the  visible date string ie mm//dd/yyyy. As a result columns are too narrow */
	void function autoSizeColumnFix(
		required numeric columnIndex /* Base-0 */
		,boolean isDateColumn=false
		,string dateMask=variables.defaultFormats[ "TIMESTAMP" ]
	){
		if( isDateColumn ){
			newWidth = estimateColumnWidth( dateMask & "00000" );
			getActiveSheet().setColumnWidth( columnIndex,newWidth );
		} else {
			getActiveSheet().autoSizeColumn( JavaCast( "int",columnIndex),true );
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

	function createRow( numeric rowNum=getNextEmptyRow(),boolean overwrite=true ){
		/* get existing row (if any)  */
		var row = getActiveSheet().getRow( JavaCast( "int",rowNum ) );
		if( overwrite AND !IsNull( row ) )
			getActiveSheet().removeRow( row ) /* forcibly remove existing row and all cells  */
		if( overwrite OR IsNull( getActiveSheet().getRow( JavaCast( "int",rowNum ) ) ) )
			row = getActiveSheet().createRow( JavaCast("int", rowNum ) );
		return row;
	}

	function createSheet( required string sheetName ){
		newSheet = workbook.createSheet( JavaCast( "String", sheetName ) );
		return newSheet;
	}

	function getActiveSheet(){
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
		var dateOnly = CreateDate( Year( dateTime ),Month(  dateTime ),Day( dateTime ) );
		if( DateCompare( value,dateOnly,"s" ) EQ 0 )
			return variables.defaultFormats.DATE;
		if( DateCompare( "1899-12-30",dateOnly,"d" ) EQ 0 )
			return variables.defaultFormats.TIME;
		return variables.defaultFormats.TIMESTAMP;
	}

	numeric function getFirstRowNum(){
		var firstRow = getActiveSheet().getFirstRowNum();
		if( firstRow EQ 0 AND getActiveSheet().getPhysicalNumberOfRows() EQ 0 )
			return -1;
		return firstRow;
	}

	numeric function getLastRowNum(){
		var lastRow = getActiveSheet().getLastRowNum();
		if( lastRow EQ 0 AND getActiveSheet().getPhysicalNumberOfRows() EQ 0 )
			return -1;//The sheet is empty. Return -1 instead of 0
		return lastRow;
	}

	numeric function getNextEmptyRow(){
		return getLastRowNum()+1;
	}

	array function getQueryColumnFormats( required query query ){
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
					//col.defaultCellStyle 	= buildCellStyle( { dataFormat = variables.defaultFormats[ col.typeName ] } );
				break;
				case "TIME":
					col.cellDataType = "TIME";
					//col.defaultCellStyle 	= buildCellStyle( { dataFormat = variables.defaultFormats[ col.typeName ] } );
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

	function initializeCell( required numeric row,required numeric column ){
		var jRow = JavaCast( "int",row-1 );
		var jColumn = JavaCast( "int",column-1 );
		var rowObject = getCellUtil().getRow( jRow,getActiveSheet() );
		var cellObject = getCellUtil().getCell( rowObject,jColumn );
		return cellObject; 
	}

	array function parseRowData( required string line,required string delimiter,boolean handleEmbeddedCommas=true ){
		var elements = ListToArray( line,delimiter );
		var potentialQuotes = 0;
		arguments.line = line.ToString();
		if( delimiter EQ "," AND handleEmbeddedCommas )
			potentialQuotes = line.ReplaceAll( "[^']", "" ).length();
		if( potentialQuotes <= 1 )
		  return elements;
		/*
			For ACF compatibility, find any values enclosed in single 
			quotes and treat them as a single element.
		*/ 
	  var currentValue = 0;
	  var nextValue = "";
		var isEmbeddedValue = false;
		var values = [];
		var buffer = CreateObject( "java.lang.StringBuilder" ).init();
		var maxElem = elements.Len();
		for( var i=1; i <= maxElem; i++ ){
		  currentValue = elements[ i ].Trim();
		  nextValue = i < maxElem ? elements[ i + 1 ] : "";
		  var isComplete = false;
		  var hasLeadingQuote = currentValue.startsWith( "'" );
		  var hasTrailingQuote = currentValue.endsWith( "'" );
		  var isFinalElem = ( i == maxElem );
			isEmbeddedValue = hasLeadingQuote;
		  isComplete	=	( isEmbeddedValue AND hasTrailingQuote );
		  // We are finished with this value if:  
		  // * no quotes were found OR
		  // * it is the final value OR
		  // * the next value is embedded in quotes
		  isComplete	=	(!isEmbeddedValue || isFinalElem || nextValue.startsWith( "'" ) );
		  if( isEmbeddedValue || isComplete ){
			  // if this a partial value, append the delimiter
			  if( isEmbeddedValue AND buffer.length() > 0 )
				  buffer.Append( "," ); 
			  buffer.Append( elements[ i ] );
		  }
		  //WriteOutput("[#i#] value=#currentValue# isEmbedded=#isEmbeddedValue# isComplete=#isComplete#"
		  //	  &" (start/end #hasLeadingQuote#/#hasTrailingQuote#) <br>");
		  if( isComplete ){
			  var finalValue = buffer.ToString();
			  var startAt = finalValue.indexOf( "'" );
			  var endAt = finalValue.lastIndexOf( "'" );
			  if( isEmbeddedValue AND startAt >= 0 AND endAt > startAt )
				  finalValue = finalValue.substring( startAt+1,endAt );
			  values.add( finalValue );
			  buffer.setLength( 0 );
			  isEmbeddedValue = false;
		  }	  
	  }
	  return values;
	}

	function setActiveSheet( string sheetName ){
		var sheetIndex = workbook.getSheetIndex( JavaCast( "string", sheetName ) ) + 1;
		workbook.setActiveSheet( JavaCast( "int",sheetIndex - 1 ) );
	}

}