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

	function createWorkBook( required string sheetName,boolean useXmlFormat=false ){
		var className = useXmlFormat? "org.apache.poi.xssf.usermodel.XSSFWorkbook": "org.apache.poi.hssf.usermodel.HSSFWorkbook";
		return loadPoi( className ).init();
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
			variables.cellUtil = loadPoi( "org.apache.poi.ss.util.CellUtil" );
		return variables.cellUtil;
	}

	function getDateUtil(){
		if( IsNull( variables.dateUtil ) )
			variables.dateUtil = loadPoi( "org.apache.poi.ss.usermodel.DateUtil" );
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

	void function flushPoiLoader(){
		lock scope="server" timeout="10"{
			StructDelete( server,"_poiLoader" );
		};
	}

	function loadPoi( required string javaclass ){
		if( !server.KeyExists( "_poiLoader" ) ){
			var paths = [];
			var libPath = ExpandPath( GetDirectoryFromPath( GetCurrentTemplatePath() ) & "lib/" );
			paths.Append( libPath & "poi-3.7-20101029.jar" );
			paths.Append( libPath & "poi-ooxml-3.7-20101029.jar" );
			paths.Append( libPath & "poi-ooxml-schemas-3.7-20101029.jar" );
			paths.Append( libPath & "dom4j-1.6.1.jar" );
			paths.Append( libPath & "geronimo-stax-api_1.0_spec-1.0.jar" );
			paths.Append( libPath & "xmlbeans-2.3.0.jar" );
			if( !server.KeyExists( "_poiLoader" ) ){
				server._poiLoader = CreateObject( "component","javaLoader.JavaLoader" ).init( loadPaths=paths,loadColdFusionClassPath=true,trustedSource=true );
			}
		}
		return server._poiLoader.create( arguments.javaclass );
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

	boolean function sheetExists( required workbook,string function sheetName,numeric sheetIndex ){
		validateSheetNameOrIndexWasProvided( argumentCollection=arguments );
		if( arguments.KeyExists( "sheetName" ) )
			arguments.sheetIndex = workbook.getSheetIndex( JavaCast("string", arguments.sheetName) ) + 1;
			//the position is valid if it an integer between 1 and the total number of sheets in the workbook
		if( sheetIndex GT 0 AND sheetIndex EQ Round( sheetIndex ) AND sheetIndex LTE workbook.getNumberOfSheets() )
			return true;
		return false;
	}

	query function sheetToQuery( required workbook,required numeric sheetIndex,numeric headerRow,boolean excludeHeaderRow=false ){
		/* Based on https://github.com/bennadel/POIUtility.cfc */
		var hasHeaderRow = arguments.KeyExists( "headerRow" );
		var poiHeaderRow = ( hasHeaderRow AND headerRow )? headerRow-1: 0;
		var columnNames=[];
		var totalColumnCount=0
		var sheet = workbook.GetSheetAt( JavaCast( "int",sheetIndex ) );
		var sheetData = [];
		var lastRowNumber = sheet.GetLastRowNum();
		// Loop over the rows in the Excel sheet.
		for( rowIndex=0; rowIndex LTE lastRowNumber; rowIndex++ ){
			var row = sheet.GetRow( JavaCast( "int",rowIndex ) );
			var rowData	=	[];
			if( IsNull( row ) )
				continue;
			var isHeaderRow = ( hasHeaderRow AND ( rowIndex EQ poiHeaderRow ) );
			var columnCount = row.GetLastCellNum();
			totalColumnCount = Max( totalColumnCount,columnCount );
			for( colIndex=0; colIndex LT columnCount; colIndex++ ){
				var cell = row.GetCell( JavaCast( "int",colIndex ) );
				if( IsNull( cell ) )
					continue;
				if( isHeaderRow ){
					try{
						var value = cell.getStringCellValue();
					}
					catch ( any exception ) {
						/* There was an error grabbing the text of the header column type. */
						var value="column#( colIndex+1 )#";
					}
					columnNames.append( value );
					if( excludeHeaderRow )
						continue;
				}
				/* When getting the value of a cell, it is important to know what type of cell value we are dealing with. If you try to grab the wrong value type, an error might be thrown. For that reason, we must check to see what type of cell we are working with. These are the cell types and they are constants of the cell object itself:
			 		
					0 - CELL_TYPE_NUMERIC
					1 - CELL_TYPE_STRING
					2 - CELL_TYPE_FORMULA
					3 - CELL_TYPE_BLANK
					4 - CELL_TYPE_BOOLEAN
					5 - CELL_TYPE_ERROR */
				
				var cellType = cell.GetCellType();
				var cellValue = "";
				/* Get the value of the cell based on the data type. The thing to worry about here is cell forumlas and cell dates. Formulas can be strange and dates are stored as numeric types. Here I will just grab dates as floats and formulas I will try to grab as numeric values. */
				if( cellType EQ cell.CELL_TYPE_NUMERIC ) {
					/* Get numeric cell data. This could be a standard number, could also be a date value. I am going to leave it up to the calling program to decide. */
					cellValue = cell.GetNumericCellValue();
				} else if( cellType EQ cell.CELL_TYPE_STRING ){
					cellValue = cell.GetStringCellValue();
				} else if( cellType EQ cell.CELL_TYPE_FORMULA ){
					/* 	Since most forumlas deal with numbers, I am going to try to grab the value as a number. If that throws an error, I will just grab it as a string value. */
					try{
						cellValue = cell.GetNumericCellValue();
					}
					catch( any exception1 ){
						// The numeric grab failed. Try to get the value as a string. If this fails, just force the empty string.
						try{
							cellValue = cell.GetStringCellValue();
						} catch( any exception2 ){
							// Force empty string.
							cellValue = "";
		 				}
					}
				} else if( cellType EQ cell.CELL_TYPE_BLANK ){
					cellValue = "";
				} else if( cellType EQ cell.CELL_TYPE_BOOLEAN ){
					cellValue = cell.GetBooleanCellValue();
				}
				rowData.append( JavaCast( "string",cellValue ) );
			}//end column loop
			if( !isHeaderRow OR !excludeHeaderRow )
				sheetData.Append( rowData );
		}//end row loop
		if( !columnNames.Len() ){
			for( var i=1; totalColumnCount; i++ ){
				columnNames.Append( "column" & i );
			}
		}
		var columnTypes = [];
		for( var columnName in columnNames ){
			columnTypes.Append( "VarChar" );
		}
		var result = QueryNew( columnNames,columnTypes,sheetData );
		return result;
	}

	void function validateSheetName( required workbook,required string sheetName ){
		if( !sheetExists( workbook=workbook,sheetName=sheetName ) )
			throw( type=exceptionType,message="Invalid Sheet Name [#arguments.SheetName#]", detail="The requested sheet was not found in the current workbook." );
	}

	void function validateSheetIndex( required workbook,required numeric sheetIndex ){
		if( !sheetExists( workbook=workbook,sheetIndex=sheetIndex ) ){
			var sheetCount = workbook.getNumberOfSheets();
			throw( type=exceptionType,message="Invalid Sheet Index [#arguments.sheetIndex#]",detail="The SheetIndex must a whole number between 1 and the total number of sheets in the workbook [#Local.sheetCount#]" );
		}
	}

	void function validateSheetNameOrIndexWasProvided( string sheetName,numeric sheetIndex ){
		if( !arguments.KeyExists( "sheetName" ) AND !arguments.KeyExists( "sheetIndex" ) )
			throw( type=exceptionType,message="Missing Required Argument", detail="Either sheetName or sheetIndex must be provided" );
		if( arguments.KeyExists( "sheetName" ) AND arguments.KeyExists( "sheetIndex" ) )
			throw( type=exceptionType,message="Too Many Arguments", detail="Only one argument is allowed. Specify either a SheetName or SheetIndex, not both" );
	}

	function workbookFromFile( required string path ){
		try{
			var inputStream 		= CreateObject( "Java","java.io.FileInputStream" ).init( path );
			var excelFileSystem = loadPoi( "org.apache.poi.poifs.filesystem.POIFSFileSystem" ).init( inputStream );
			var workbook 				= loadPoi( "org.apache.poi.hssf.usermodel.HSSFWorkbook" ).init( excelFileSystem );
		}
		finally{
			inputStream.close();
		}
		return workbook;
	}

}
