<cfscript>
private void function addInfoBinary( required workbook,required struct info ){
	workbook.createInformationProperties();
	var documentSummaryInfo = workbook.getDocumentSummaryInformation();
	var summaryInfo = workbook.getSummaryInformation();	
	
	for( var key in info ) {
		switch( key ){
			case "author": 				
				summaryInfo.setAuthor( JavaCast( "string",info.author ) );
				break;
			case "category":
				documentSummaryInfo.setCategory( JavaCast( "string",info.category ) );
				break;
			case "lastauthor":
				summaryInfo.setLastAuthor( JavaCast( "string",info.lastauthor ) );
				break;
			case "comments":
				summaryInfo.setComments( JavaCast( "string",info.comments ) );	
				break;
			case "keywords":
				summaryInfo.setKeywords( JavaCast( "string",info.keywords ) );
				break;
			case "manager":
				documentSummaryInfo.setManager( JavaCast( "string",info.manager ) );
				break;
			case "company":
				documentSummaryInfo.setCompany( JavaCast( "string",info.company ) );
				break;
			case "subject":
				summaryInfo.setSubject( JavaCast( "string",info.subject ) );
				break;
			case "title":
				summaryInfo.setTitle( JavaCast( "string",info.title ) );
				break;
		}
	}	
}

private void function addInfoXml( required workbook,required struct info ){
	var documentProperties = workbook.getProperties().getExtendedProperties().getUnderlyingProperties();
	var coreProperties = workbook.getProperties().getCoreProperties();
	for (var key in info ) {
		switch( key ){
			case "author": 				
				coreProperties.setCreator( JavaCast( "string",info[ key ] )  );
				break;
			case "category": 				
				coreProperties.setCategory( JavaCast( "string",info[ key ] ) );
				break;
			case "lastauthor": 
				// TODO: This does not seem to be working. Not sure why
				coreProperties.getUnderlyingProperties().setLastModifiedByProperty( JavaCast( "string",info[ key ] ) );
				break;
			case "comments": 				
				coreProperties.setDescription( JavaCast( "string",info[ key ] ) );	
				break;
			case "keywords": 				
				coreProperties.setKeywords( JavaCast( "string",info[ key ] ) );
				break;
			case "subject": 				
				coreProperties.setSubjectProperty( JavaCast( "string",info[ key ] ) );
				break;
			case "title": 				
				coreProperties.setTitle( JavaCast( "string",info[ key ] ) );
				break;
			case "manager": 	
				documentProperties.setManager( JavaCast( "string",info[ key ] ) );
				break;
			case "company": 				
				documentProperties.setCompany( JavaCast( "string",info[ key ] ) );
				break;
		}
	}
}

/* Workaround for an issue with autoSizeColumn(). It does not seem to handle 
	date cells properly. It measures the length of the date "number", instead of 
	the  visible date string ie mm//dd/yyyy. As a result columns are too narrow */
private void function autoSizeColumnFix(
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

private boolean function cellExists( required workbook,required numeric rowNumber,required numeric columnNumber ){
	var rowIndex = rowNumber-1;
	var columnIndex = columnNumber=1;
	var checkRow = this.getActiveSheet( workbook ).getRow( JavaCast( "int",rowIndex ) );
	return !IsNull( checkRow ) AND !IsNull( checkRow.getCell( JavaCast( "int",columnIndex ) ) );
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

private void function deleteSheetAtIndex( required workbook,required numeric sheetIndex ){
	workbook.removeSheetAt( JavaCast( "int",sheetIndex ) );
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

private function getCellUtil(){
	if( IsNull( variables.cellUtil ) )
		variables.cellUtil = loadPoi( "org.apache.poi.ss.util.CellUtil" );
	return variables.cellUtil;
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
		return variables.defaultFormats.DATE;
	if( DateCompare( "1899-12-30",dateOnly,"d" ) EQ 0 )
		return variables.defaultFormats.TIME;
	return variables.defaultFormats.TIMESTAMP;
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
				col.defaultCellStyle 	= this.buildCellStyle( workbook,{ dataFormat = variables.defaultFormats[ col.typeName ] } );
			break;
			case "TIME":
				col.cellDataType = "TIME";
				col.defaultCellStyle 	= this.buildCellStyle( workbook,{ dataFormat = variables.defaultFormats[ col.typeName ] } );
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

private function loadPoi( required string javaclass ){
	if( !server.KeyExists( "_poiLoader" ) ){
		var paths = [];
		var libPath = ExpandPath( GetDirectoryFromPath( GetCurrentTemplatePath() ) & "lib/" );
		paths.Append( libPath & "poi-3.11-20141221.jar" );
		paths.Append( libPath & "poi-ooxml-3.11-20141221.jar" );
		paths.Append( libPath & "poi-ooxml-schemas-3.11-20141221.jar" );
		paths.Append( libPath & "xmlbeans-2.6.0.jar" );
		if( !server.KeyExists( "_poiLoader" ) ){
			server._poiLoader = CreateObject( "component","javaLoader.JavaLoader" ).init( loadPaths=paths,loadColdFusionClassPath=true,trustedSource=true );
		}
	}
	return server._poiLoader.create( arguments.javaclass );
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
	/*
		For ACF compatibility, find any values enclosed in single 
		quotes and treat them as a single element.
	*/ 
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

private boolean function sheetExists( required workbook,string sheetName,numeric sheetNumber ){
	validateSheetNameOrNumberWasProvided( argumentCollection=arguments );
	if( arguments.KeyExists( "sheetName" ) )
		arguments.sheetNumber = this.getSheetIndexFromName( workbook,sheetName )+1;
		//the position is valid if it an integer between 1 and the total number of sheets in the workbook
	if( sheetNumber AND ( sheetNumber EQ Round( sheetNumber ) ) AND ( sheetNumber LTE workbook.getNumberOfSheets() ) )
		return true;
	return false;
}

private query function sheetToQuery( required workbook,string sheetName,numeric sheetNumber=1,numeric headerRow,boolean includeHeaderRow=false,boolean includeBlankRows=false ){
	/* Based on https://github.com/bennadel/POIUtility.cfc */
	var hasHeaderRow = arguments.KeyExists( "headerRow" );
	var poiHeaderRow = ( hasHeaderRow AND headerRow )? headerRow-1: 0;
	var columnNames=[];
	var totalColumnCount=0
	if( arguments.KeyExists( "sheetName" ) ){
		validateSheetExistsWithName( workbook,sheetName );
		arguments.sheetNumber = getSheetIndexFromName( workbook,sheetName )+1;
	}
	var sheet = workbook.GetSheetAt( JavaCast( "int",sheetNumber-1 ) );
	var sheetData = [];
	var lastRowNumber = sheet.GetLastRowNum();
	// Loop over the rows in the Excel sheet.
	for( rowIndex=0; rowIndex LTE lastRowNumber; rowIndex++ ){
		var isHeaderRow = ( hasHeaderRow AND ( rowIndex EQ poiHeaderRow ) );
		var rowData	=	[];
		var row = sheet.GetRow( JavaCast( "int",rowIndex ) );
		if( IsNull( row ) ){
			if( !isHeaderRow OR includeHeaderRow )
				sheetData.Append( rowData );
			continue;
		}
		var columnCount = row.GetLastCellNum();
		totalColumnCount = Max( totalColumnCount,columnCount );
		for( colIndex=0; colIndex LT columnCount; colIndex++ ){
			var cell = row.GetCell( JavaCast( "int",colIndex ) );
			if( IsNull( cell ) ){
				if( includeBlankRows )
					rowData.append( JavaCast( "string","" ) );
				continue;
			}
			if( isHeaderRow ){
				try{
					var value = cell.getStringCellValue();
				}
				catch ( any exception ) {
					/* There was an error grabbing the text of the header column type. */
					var value="column#( colIndex+1 )#";
				}
				columnNames.append( value );
				if( !includeHeaderRow )
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
		if( !isHeaderRow OR includeHeaderRow )
			sheetData.Append( rowData );
	}//end row loop
	if( !columnNames.Len() ){
		for( var i=1; i LTE totalColumnCount; i++ ){
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

private void function validateSheetExistsWithName( required workbook,required string sheetName ){
	if( !this.sheetExists( workbook=workbook,sheetName=sheetName ) )
		throw( type=exceptionType,message="Invalid sheet name [#sheetName#]", detail="The specified sheet was not found in the current workbook." );
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
		var file = CreateObject( "java","java.io.FileInputStream" ).init( path );
		var workbook = loadPoi( "org.apache.poi.ss.usermodel.WorkbookFactory" ).create( file );
		return workbook;
	}
	finally{
		file.close();
	}
}
</cfscript>