component{

	variables.version = "2.7.0-develop";
	variables.javaLoaderName = "spreadsheetLibraryClassLoader-#variables.version#-#Hash( GetCurrentTemplatePath() )#";
	variables.javaLoaderDotPath = "javaLoader.JavaLoader";
	variables.dateFormats = {
		DATE: "yyyy-mm-dd"
		,DATETIME: "yyyy-mm-dd HH:nn:ss"
		,TIME: "hh:mm:ss"
		,TIMESTAMP: "yyyy-mm-dd hh:mm:ss"
	};
	variables.exceptionType = "cfsimplicity.lucee.spreadsheet";
	variables.isACF = ( server.coldfusion.productname IS "ColdFusion Server" );
	variables.javaClassesLastLoadedVia = "Nothing loaded yet";
	variables.engineSupportsWriteEncryption = !isACF;

	variables.HSSFWorkbookClassName = "org.apache.poi.hssf.usermodel.HSSFWorkbook";
	variables.XSSFWorkbookClassName = "org.apache.poi.xssf.usermodel.XSSFWorkbook";
	variables.SXSSFWorkbookClassName = "org.apache.poi.xssf.streaming.SXSSFWorkbook";

	function init( struct dateFormats, string javaLoaderDotPath, boolean requiresJavaLoader=true ){
		if( arguments.KeyExists( "dateFormats" ) )
			overrideDefaultDateFormats( arguments.dateFormats );
		if( arguments.KeyExists( "javaLoaderDotPath" ) ) // Option to use the dot path of an existing javaloader installation to save duplication
			variables.javaLoaderDotPath = arguments.javaLoaderDotPath;
		variables.requiresJavaLoader = arguments.requiresJavaLoader;
		return this;
	}

	/* Meta utilities */

	private void function overrideDefaultDateFormats( required struct formats ){
		for( var format in arguments.formats ){
			if( !variables.dateFormats.KeyExists( format ) )
				Throw( type=exceptionType, message="Invalid date format key", detail="'#format#' is not a valid dateformat key. Valid keys are DATE, DATETIME, TIME and TIMESTAMP" );
			variables.dateFormats[ format ] = arguments.formats[ format ];
		}
	}

	public void function flushPoiLoader(){
		lock scope="server" timeout="10" {
			StructDelete( server, javaLoaderName );
		};
	}

	public struct function getDateFormats(){
		return dateFormats;
	}

	public struct function getEnvironment(){
		return {
			dateFormats: dateFormats
			,engine: server.coldfusion.productname & " " & ( isACF? server.coldfusion.productversion: ( server.lucee.version?: "?" ) )
			,engineSupportsEncryption: engineSupportsWriteEncryption //for backwards compat only //TODO remove on next major version
			,engineSupportsWriteEncryption: engineSupportsWriteEncryption
			,javaLoaderDotPath: javaLoaderDotPath
			,javaClassesLastLoadedVia: javaClassesLastLoadedVia
			,javaLoaderName: javaLoaderName
			,requiresJavaLoader: requiresJavaLoader
			,version: version
		};
	}

	// Diagnostic tool: check physical path of a specific class
	public void function dumpPathToClass( required string className ){
		var classLoader = loadClass( arguments.className ).getClass().getClassLoader();
		var path = classLoader.getResource( arguments.className.Replace( ".", "/", "all" ) & ".class" ).getPath();
		WriteDump( path );
	}

	/* MAIN PUBLIC API */

	/* Convenenience */

	public binary function binaryFromQuery(
		required query data
		,boolean addHeaderRow=true
		,boolean boldHeaderRow=true
		,boolean xmlFormat=false
		,boolean streamingXml=false
		,numeric streamingWindowSize=100
	){
		/* Pass in a query and get a spreadsheet binary file ready to stream to the browser */
		var workbook = workbookFromQuery( argumentCollection=arguments );
		var binary = readBinary( workbook );
		cleanUpStreamingXml( workbook );
		return binary;
	}

	public function csvToQuery(
		string csv=""
		,string filepath=""
		,boolean firstRowIsHeader=false
		,boolean trim=true
		,string delimiter
	){
		var csvIsString = arguments.csv.Len();
		var csvIsFile = arguments.filepath.Len();
		if( !csvIsString AND !csvIsFile )
			Throw( type=exceptionType, message="Missing required argument", detail="Please provide either a csv string (csv), or the path of a file containing one (filepath)." );
		if( csvIsString AND csvIsFile )
			Throw( type=exceptionType, message="Mutually exclusive arguments: 'csv' and 'filepath'", detail="Only one of either 'filepath' or 'csv' arguments may be provided." );
		if(	csvIsFile ){
			if( !FileExists( arguments.filepath ) )
				Throw( type=exceptionType, message="Non-existant file", detail="Cannot find a file at #arguments.filepath#" );
			if( !isCsvOrTextFile( arguments.filepath ) )
				Throw( type=exceptionType, message="Invalid csv file", detail="#arguments.filepath# does not appear to be a text/csv file" );
			arguments.csv = FileRead( arguments.filepath );
		}
		var format = loadClass( "org.apache.commons.csv.CSVFormat" )[ JavaCast( "string", "RFC4180" ) ];
		format = format.withIgnoreSurroundingSpaces();//stop spaces between fields causing problems with embedded lines
		if( arguments.trim )
			arguments.csv = arguments.csv.Trim();
		if( arguments.KeyExists( "delimiter" ) )
			format = format.withDelimiter( JavaCast( "string", arguments.delimiter ) );
		var parsed = loadClass( "org.apache.commons.csv.CSVParser" ).parse( arguments.csv, format );
		var records = parsed.getRecords();
		var rows = [];
		var maxColumnCount = 0;
		for( var record in records ){
			var row = [];
			var columnNumber = 0;
			var iterator = record.iterator();
			while( iterator.hasNext() ){
				columnNumber++;
				maxColumnCount = Max( maxColumnCount, columnNumber );
				row.Append( iterator.next() );
			}
			rows.Append( row );
		}
		var columnList = [];
		if( arguments.firstRowIsHeader )
			var headerRow = rows[ 1 ];
		for( var i=1; i LTE maxColumnCount; i++ ){
			if( arguments.firstRowIsHeader AND !IsNull( headerRow[ i ] ) AND headerRow[ i ].Len() ){
				columnList.Append( JavaCast( "string", headerRow[ i ] ) );
				continue;
			}
			columnList.Append( "column#i#" );
		}
		if( arguments.firstRowIsHeader )
			rows.DeleteAt( 1 );
		return _queryNew( columnList, "", rows );
	}

	public void function download( required workbook, required string filename, string contentType ){
		var safeFilename = filenameSafe( arguments.filename );
		var filenameWithoutExtension = safeFilename.REReplace( "\.xlsx?$", "" );
		var extension = isXmlFormat( arguments.workbook )? "xlsx": "xls";
		arguments.filename = filenameWithoutExtension & "." & extension;
		var binary = readBinary( arguments.workbook );
		cleanUpStreamingXml( arguments.workbook );
		if( !arguments.KeyExists( "contentType" ) )
			arguments.contentType = isXmlFormat( arguments.workbook )? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "application/msexcel";
		downloadBinaryVariable( binary, arguments.filename, arguments.contentType );
	}

	public void function downloadFileFromQuery(
		required query data
		,required string filename
		,boolean addHeaderRow=true
		,boolean boldHeaderRow=true
		,boolean xmlFormat=false
		,string contentType
		,boolean streamingXml=false
		,numeric streamingWindowSize=100
	){
		var safeFilename = filenameSafe( arguments.filename );
		var filenameWithoutExtension = safeFilename.REReplace( "\.xlsx?$","" );
		var extension = ( arguments.xmlFormat || arguments.streamingXml )? "xlsx": "xls";
		arguments.filename = filenameWithoutExtension & "." & extension;
		var binary = binaryFromQuery( arguments.data, arguments.addHeaderRow, arguments.boldHeaderRow, arguments.xmlFormat, arguments.streamingXml, arguments.streamingWindowSize );
		if( !arguments.KeyExists( "contentType" ) )
			arguments.contentType = arguments.xmlFormat? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "application/msexcel";
		downloadBinaryVariable( binary, arguments.filename, arguments.contentType );
	}

	public void function downloadCsvFromFile(
		required string src
		,required string filename
		,string contentType="text/csv"
		,string columns
		,string columnNames
		,numeric headerRow
		,string rows
		,string sheetName
		,numeric sheetNumber // 1-based
		,boolean includeHeaderRow=false
		,boolean includeBlankRows=false
		,boolean fillMergedCellsWithVisibleValue=false
	){
		arguments.format = "csv";
		var csv = read( argumentCollection=arguments );
		var binary = ToBinary( ToBase64( csv.Trim() ) );
		var safeFilename = filenameSafe( arguments.filename );
		var filenameWithoutExtension = safeFilename.REReplace( "\.csv$","" );
		var extension = "csv";
		arguments.filename = filenameWithoutExtension & "." & extension;
		downloadBinaryVariable( binary, arguments.filename, arguments.contentType );
	}

	public any function workbookFromCsv(
		string csv
		,string filepath
		,boolean firstRowIsHeader=false
		,boolean boldHeaderRow=true
		,boolean trim=true
		,boolean xmlFormat=false
		,string delimiter
	){
		var conversionArgs = {
			firstRowIsHeader: arguments.firstRowIsHeader
			,trim: arguments.trim
		};
		if( arguments.KeyExists( "csv" ) )
			conversionArgs.csv = arguments.csv;
		if( arguments.KeyExists( "filepath" ) )
			conversionArgs.filepath = arguments.filepath;
		if( arguments.KeyExists( "delimiter" ) )
			conversionArgs.delimiter = arguments.delimiter;
		var data = csvToQuery( argumentCollection=conversionArgs );
		return workbookFromQuery(
			data=data
			,addHeaderRow=arguments.firstRowIsHeader
			,boldHeaderRow=arguments.boldHeaderRow
			,xmlFormat=arguments.xmlFormat
		);
	}

	public any function workbookFromQuery(
		required query data
		,boolean addHeaderRow=true
		,boolean boldHeaderRow=true
		,boolean xmlFormat=false
		,boolean streamingXml=false
		,numeric streamingWindowSize=100
	){
		var workbook = new( xmlFormat=arguments.xmlFormat, streamingXml=arguments.streamingXml, streamingWindowSize=arguments.streamingWindowSize );
		if( arguments.addHeaderRow ){
			var columns = _queryColumnArray( arguments.data );
			addRow( workbook, columns );
			if( arguments.boldHeaderRow )
				formatRow( workbook, { bold: true }, 1 );
			addRows( workbook, arguments.data, 2, 1 );
		}
		else
			addRows( workbook, arguments.data );
		return workbook;
	}

	public void function writeFileFromQuery(
		required query data
		,required string filepath
		,boolean overwrite=false
		,boolean addHeaderRow=true
		,boldHeaderRow=true
		,xmlFormat=false
		,boolean streamingXml=false
		,numeric streamingWindowSize=100
	){
		if( !arguments.xmlFormat AND ( ListLast( arguments.filepath, "." ) IS "xlsx" ) )
			arguments.xmlFormat = true;
		var workbook = workbookFromQuery(
			data=arguments.data
			,addHeaderRow=arguments.addHeaderRow
			,boldHeaderRow=arguments.boldHeaderRow
			,xmlFormat=arguments.xmlFormat
			,streamingXml=arguments.streamingXml
			,streamingWindowSize=arguments.streamingWindowSize
		);
		if( xmlFormat AND ( ListLast( arguments.filepath, "." ) IS "xls" ) )
			arguments.filepath &= "x";// force to .xlsx
		write( workbook=workbook, filepath=arguments.filepath, overwrite=arguments.overwrite );
	}

	/* End convenience methods */

	public void function addAutofilter( required workbook, required string cellRange ){
		arguments.cellRange = arguments.cellRange.Trim();
		if( arguments.cellRange.IsEmpty() )
			Throw( type=exceptionType, message="Empty cellRange argument", detail="You must provide a cell range reference in the form 'A1:Z1'" );
		getActiveSheet( arguments.workbook ).setAutoFilter( getCellRangeAddressFromReference( arguments.cellRange ) );
	}

	public void function addColumn(
		required workbook
		,required string data /* Delimited list of cell values */
		,numeric startRow
		,numeric startColumn
		,boolean insert=true
		,string delimiter=","
		,boolean autoSize=false
	){
		var row = 0;
		var cell = 0;
		var oldCell = 0;
		var rowNum = ( arguments.KeyExists( "startRow" ) AND arguments.startRow )? arguments.startRow -1: 0;
		var cellNum = 0;
		var lastCellNum = 0;
		var cellValue = 0;
		var sheet = getActiveSheet( arguments.workbook );
		if( arguments.KeyExists( "startColumn" ) )
			cellNum = ( arguments.startColumn -1 );
		else {
			row = sheet.getRow( rowNum );
			/* if this row exists, find the next empty cell number. note: getLastCellNum() returns the cell index PLUS ONE or -1 if not found */
			if( !IsNull( row ) AND row.getLastCellNum() GT 0 )
				cellNum = row.getLastCellNum();
			else
				cellNum = 0;
		}
		var columnNumber = ( cellNum +1 );
		var columnData = ListToArray( arguments.data, arguments.delimiter );
		for( var cellValue in columnData ){
			/* if rowNum is greater than the last row of the sheet, need to create a new row  */
			if( rowNum GT sheet.getLastRowNum() OR IsNull( sheet.getRow( rowNum ) ) )
				row = createRow( arguments.workbook, rowNum );
			else
				row = sheet.getRow( rowNum );
			/* POI doesn't have any 'shift column' functionality akin to shiftRows() so inserts get interesting */
			/* ** Note: row.getLastCellNum() returns the cell index PLUS ONE or -1 if not found */
			if( arguments.insert AND ( cellNum LT row.getLastCellNum() ) ){
				/*  need to get the last populated column number in the row, figure out which cells are impacted, and shift the impacted cells to the right to make room for the new data */
				lastCellNum = row.getLastCellNum();
				for( var i = lastCellNum; i EQ cellNum; i-- ){
					oldCell = row.getCell( JavaCast( "int", i -1 ) );
					if( !IsNull( oldCell ) ){
						cell = createCell( row, i );
						cell.setCellStyle( oldCell.getCellStyle() );
						var cellValue = getCellValueAsType( arguments.workbook, oldCell );
						setCellValueAsType( arguments.workbook, oldCell, cellValue );
						cell.setCellComment( oldCell.getCellComment() );
					}
				}
			}
			cell = createCell( row,cellNum );
			setCellValueAsType( arguments.workbook, cell, cellValue );
			rowNum++;
		}
		if( arguments.autoSize )
			autoSizeColumn( arguments.workbook, columnNumber );
	}

	public void function addFreezePane(
		required workbook
		,required numeric freezeColumn
		,required numeric freezeRow
		,numeric leftmostColumn //left column visible in right pane
		,numeric topRow //top row visible in bottom pane
	){
		var sheet = getActiveSheet( arguments.workbook );
		if( arguments.KeyExists( "leftmostColumn" ) AND !arguments.KeyExists( "topRow" ) )
			arguments.topRow = arguments.freezeRow;
		if( arguments.KeyExists( "topRow" ) AND !arguments.KeyExists( "leftmostColumn" ) )
			arguments.leftmostColumn = arguments.freezeColumn;
		/* createFreezePane() operates on the logical row/column numbers as opposed to physical, so no need for n-1 stuff here */
		if( !arguments.KeyExists( "leftmostColumn" ) ){
			sheet.createFreezePane( JavaCast( "int", arguments.freezeColumn ), JavaCast( "int", arguments.freezeRow ) );
			return;
		}
		// POI lets you specify an active pane if you use createSplitPane() here
		sheet.createFreezePane(
			JavaCast( "int", arguments.freezeColumn )
			,JavaCast( "int", arguments.freezeRow )
			,JavaCast( "int", arguments.leftmostColumn )
			,JavaCast( "int", arguments.topRow )
		);
	}

	public void function addImage(
		required workbook
		,string filepath
		,imageData
		,string imageType
		,required string anchor
	){
		/*
			TODO: Should we allow for passing in of a boolean indicating whether or not an image resize should happen (only works on jpg and png)? Currently does not resize. If resize is performed, it does mess up passing in x/y coordinates for image positioning.
		 */
		if( !arguments.KeyExists( "filepath" ) AND !arguments.KeyExists( "imageData" ) )
			Throw( type=exceptionType, message="Invalid argument combination", detail="You must provide either a file path or an image object" );
		if( arguments.KeyExists( "imageData" ) AND !arguments.KeyExists( "imageType" ) )
			Throw( type=exceptionType, message="Invalid argument combination", detail="If you specify an image object, you must also provide the imageType argument" );
		var numberOfAnchorElements = ListLen( arguments.anchor );
		if( ( numberOfAnchorElements NEQ 4 ) AND ( numberOfAnchorElements NEQ 8 ) )
			Throw( type=exceptionType, message="Invalid anchor argument", detail="The anchor argument must be a comma-delimited list of integers with either 4 or 8 elements" );
		//we'll need the image type int in all cases
		if( arguments.KeyExists( "filepath" ) ){
			if( !FileExists( arguments.filepath ) )
				Throw( type=exceptionType, message="Non-existent file", detail="The specified file does not exist." );
			try{
				arguments.imageType = ListLast( FileGetMimeType( arguments.filepath ), "/" );
			}
			catch( any exception ){
				Throw( type=exceptionType, message="Could Not Determine Image Type", detail="An image type could not be determined from the filepath provided" );
			}
		}
		else if( !arguments.KeyExists( "imageType" ) )
			Throw( type=exceptionType, message="Could Not Determine Image Type", detail="An image type could not be determined from the filepath or imagetype provided" );
		arguments.imageType	=	arguments.imageType.UCase();
		switch( arguments.imageType ){
			case "DIB": case "EMF": case "JPEG": case "PICT": case "PNG": case "WMF":
				var imageTypeIndex = arguments.workbook[ "PICTURE_TYPE_" & arguments.imageType ];
			break;
			case "JPG":
				var imageTypeIndex = arguments.workbook.PICTURE_TYPE_JPEG;
			break;
			default:
				Throw( type=exceptionType, message="Invalid Image Type", detail="Valid image types are DIB, EMF, JPG, JPEG, PICT, PNG, and WMF" );
		}
		if( arguments.KeyExists( "filepath" ) ){
			try{
				var inputStream = CreateObject( "java", "java.io.FileInputStream" ).init( JavaCast( "string", arguments.filepath ) );
				var ioUtils = loadClass( "org.apache.poi.util.IOUtils" );
				var bytes = ioUtils.toByteArray( inputStream );
			}
			finally{
				if( local.KeyExists( "inputStream" ) )
					inputStream.close();
			}
		}
		else
			var bytes = ToBinary( arguments.imageData );
		var imageIndex = arguments.workbook.addPicture( bytes, JavaCast( "int", imageTypeIndex ) );
		var clientAnchorClass = isXmlFormat( arguments.workbook )
				? "org.apache.poi.xssf.usermodel.XSSFClientAnchor"
				: "org.apache.poi.hssf.usermodel.HSSFClientAnchor";
		var theAnchor = loadClass( clientAnchorClass ).init();
		if( numberOfAnchorElements EQ 4 ){
			theAnchor.setRow1( JavaCast( "int", ListFirst( arguments.anchor ) -1 ) );
			theAnchor.setCol1( JavaCast( "int", ListGetAt( arguments.anchor, 2 ) -1 ) );
			theAnchor.setRow2( JavaCast( "int", ListGetAt( arguments.anchor, 3 ) -1 ) );
			theAnchor.setCol2( JavaCast( "int", ListLast( arguments.anchor ) -1 ) );
		}
		else if( numberOfAnchorElements EQ 8 ){
			theAnchor.setDx1( JavaCast( "int", ListFirst( arguments.anchor ) ) );
			theAnchor.setDy1( JavaCast( "int", ListGetAt( arguments.anchor, 2 ) ) );
			theAnchor.setDx2( JavaCast( "int", ListGetAt( arguments.anchor, 3 ) ) );
			theAnchor.setDy2( JavaCast( "int", ListGetAt( arguments.anchor, 4 ) ) );
			theAnchor.setRow1( JavaCast( "int", ListGetAt( arguments.anchor, 5 ) -1 ) );
			theAnchor.setCol1( JavaCast( "int", ListGetAt( arguments.anchor, 6 ) -1 ) );
			theAnchor.setRow2( JavaCast( "int", ListGetAt( arguments.anchor, 7 ) -1 ) );
			theAnchor.setCol2( JavaCast( "int", ListLast( arguments.anchor ) -1 ) );
		}
		/* TODO: need to look into createDrawingPatriarch() vs. getDrawingPatriarch() since create will kill any existing images. getDrawingPatriarch() throws  a null pointer exception when an attempt is made to add a second image to the spreadsheet  */
		var drawingPatriarch = getActiveSheet( arguments.workbook ).createDrawingPatriarch();
		var picture = drawingPatriarch.createPicture( theAnchor, imageIndex );
		/* Disabling this for now--maybe let people pass in a boolean indicating whether or not they want the image resized?
		 if this is a png or jpg, resize the picture to its original size (this doesn't work for formats other than jpg and png)
			<cfif imgTypeIndex eq getWorkbook().PICTURE_TYPE_JPEG or imgTypeIndex eq getWorkbook().PICTURE_TYPE_PNG>
				<cfset picture.resize() />
			</cfif>
		*/
	}

	public void function addInfo( required workbook, required struct info ){
		/* Valid struct keys are author, category, lastauthor, comments, keywords, manager, company, subject, title */
		if( isBinaryFormat( arguments.workbook ) )
			addInfoBinary( arguments.workbook, arguments.info );
		else
			addInfoXml( arguments.workbook, arguments.info );
	}

	public void function addPageBreaks( required workbook, string rowBreaks="", string columnBreaks="" ){
		arguments.rowBreaks = Trim( arguments.rowBreaks ); //Dont' use member function in case value is in fact numeric
		arguments.columnBreaks = Trim( columnBreaks );
		if( arguments.rowBreaks.IsEmpty() AND arguments.columnBreaks.IsEmpty() )
			Throw( type=exceptionType, message="Missing argument", detail="You must specify the rows and/or columns at which page breaks should be added." );
		arguments.rowBreaks = arguments.rowBreaks.ListToArray();
		arguments.columnBreaks = arguments.columnBreaks.ListToArray();
		var sheet = getActiveSheet( arguments.workbook );
		sheet.setAutoBreaks( false ); // Not sure if this is necessary: https://stackoverflow.com/a/14900320/204620
		for( var rowNumber in arguments.rowBreaks )
			sheet.setRowBreak( JavaCast( "int", ( rowNumber -1 ) ) );
		for( var columnNumber in arguments.columnBreaks )
			sheet.setcolumnBreak( JavaCast( "int", ( columnNumber -1 ) ) );
	}

	public void function addPrintGridlines( required workbook ){
		getActiveSheet( arguments.workbook ).setPrintGridlines( JavaCast( "boolean", true ) );
	}

	public void function addRow(
		required workbook
		,required data /* Delimited list of data, OR array */
		,numeric row
		,numeric column=1
		,boolean insert=true
		,string delimiter=","
		,boolean handleEmbeddedCommas=true /* When true, values enclosed in single quotes are treated as a single element like in ACF. Only applies when the delimiter is a comma. */
		,boolean autoSizeColumns=false
	){
		if( arguments.KeyExists( "row" ) AND ( arguments.row LTE 0 ) )
			Throw( type=exceptionType, message="Invalid row value", detail="The value for row must be greater than or equal to 1." );
		if( arguments.KeyExists( "column" ) AND ( arguments.column LTE 0 ) )
			Throw( type=exceptionType, message="Invalid column value", detail="The value for column must be greater than or equal to 1." );
		if( !arguments.insert AND !arguments.KeyExists( "row") )
			Throw( type=exceptionType, message="Missing row value", detail="To replace a row using 'insert', please specify the row to replace." );
		var lastRow = getNextEmptyRow( arguments.workbook );
		//If the requested row already exists...
		if( arguments.KeyExists( "row" ) AND ( arguments.row LTE lastRow ) ){
			if( arguments.insert )
				shiftRows( arguments.workbook, arguments.row, lastRow, 1 );//shift the existing rows down (by one row)
			else
				deleteRow( arguments.workbook, arguments.row );//otherwise, clear the entire row
		}
		var theRow = arguments.KeyExists( "row" )? createRow( arguments.workbook, arguments.row -1 ): createRow( arguments.workbook );
		var dataIsArray = IsArray( arguments.data );
		var rowValues = dataIsArray? arguments.data: parseRowData( arguments.data, arguments.delimiter, arguments.handleEmbeddedCommas );
		var cellIndex = arguments.column -1;
		for( var cellValue in rowValues ){
			var cell = createCell( theRow, cellIndex );
			setCellValueAsType( arguments.workbook, cell, Trim( cellValue ) );
			if( arguments.autoSizeColumns )
				autoSizeColumn( arguments.workbook, arguments.column );
			cellIndex++;
		}
	}

	public void function addRows(
		required workbook
		,required data // query or array
		,numeric row
		,numeric column=1
		,boolean insert=true
		,boolean autoSizeColumns=false
		,boolean includeQueryColumnNames=false
	){
		var dataIsQuery = IsQuery( arguments.data );
		var dataIsArray = IsArray( arguments.data );
		if( !dataIsQuery && !dataIsArray )
			Throw( type=exceptionType, message="Invalid data argument", detail="The data passed in must be either a query or an array of row arrays." );
		var totalRows = dataIsQuery? arguments.data.recordCount: arguments.data.Len();
		if( totalRows == 0 )
			return;
		// array data must be an array of arrays, not structs
		if( dataIsArray && !IsArray( arguments.data[ 1 ] ) )
			Throw( type=exceptionType, message="Invalid data argument", detail="Data passed as an array must be an array of arrays, one per row" );
		var lastRow = getNextEmptyRow( arguments.workbook );
		var insertAtRowIndex = arguments.KeyExists( "row" )? arguments.row -1: getNextEmptyRow( arguments.workbook );
		if( arguments.KeyExists( "row" ) AND ( arguments.row LTE lastRow ) AND arguments.insert )
			shiftRows( arguments.workbook, arguments.row, lastRow, totalRows );
		var currentRowIndex = insertAtRowIndex;
		var dateUtil = getDateUtil();
		if( dataIsQuery ){
			var queryColumns = getQueryColumnFormats( arguments.data );
			var cellIndex = ( arguments.column -1 );
			if( arguments.includeQueryColumnNames ){
				var columnNames = _queryColumnArray( arguments.data );
				addRow( workbook=arguments.workbook, data=columnNames, row=currentRowIndex +1, column=arguments.column );
				currentRowIndex++;
			}
			for( var dataRow in arguments.data ){
				var newRow = createRow( arguments.workbook, currentRowIndex, false );
				cellIndex = ( arguments.column -1 );//reset for this row
	   		/* populate all columns in the row */
	   		for( var queryColumn in queryColumns ){
	   			var cell = createCell( newRow, cellIndex, false );
					var value = dataRow[ queryColumn.name ];
					/* Cast the values to the correct type  */
					switch( queryColumn.cellDataType ){
						case "DOUBLE":
							setCellValueAsType( arguments.workbook, cell, value, "numeric" );
							break;
						case "DATE":
						case "TIME":
							setCellValueAsType( arguments.workbook, cell, value, "date" );
							break;
						case "BOOLEAN":
							setCellValueAsType( arguments.workbook, cell, value, "boolean" );
							break;
						default:
							if( IsSimpleValue( value ) AND !Len( value ) ) //NB don't use member function: won't work if numeric
								setCellValueAsType( arguments.workbook, cell, value, "blank" );
							else
								setCellValueAsType( arguments.workbook, cell, value, "string" );
					}
					cellIndex++;
	   		}
	   		currentRowIndex++;
			}
			if( arguments.autoSizeColumns ){
				var numberOfColumns = queryColumns.Len();
				var thisColumn = arguments.column;
				for( var i = thisColumn; i LTE numberOfColumns; i++ ){
					autoSizeColumn( arguments.workbook, thisColumn );
					thisColumn++;
				}
			}
		}
		else { //data is an array
			for( var dataRow in arguments.data ){
				var newRow = createRow( arguments.workbook, currentRowIndex, false );
				var cellIndex = ( arguments.column -1 );
	   		/* populate all columns in the row */
	   		for( var cellValue in dataRow ){
					var cell = createCell( newRow, cellIndex );
					setCellValueAsType( arguments.workbook, cell, Trim( cellValue ) );
					if( arguments.autoSizeColumns )
						autoSizeColumn( arguments.workbook, arguments.column );
					cellIndex++;
				}
				currentRowIndex++;
	   	}
		}
	}

	public void function addSplitPane(
		required workbook
		,required numeric xSplitPosition
		,required numeric ySplitPosition
		,required numeric leftmostColumn
		,required numeric topRow
		,string activePane="UPPER_LEFT" //Valid values are LOWER_LEFT, LOWER_RIGHT, UPPER_LEFT, and UPPER_RIGHT
	){
		var sheet = getActiveSheet( arguments.workbook );
		arguments.activePane = activeSheet[ "PANE_#arguments.activePane#" ];
		sheet.createSplitPane(
			JavaCast( "int", arguments.xSplitPosition )
			,JavaCast( "int", arguments.ySplitPosition )
			,JavaCast( "int", arguments.leftmostColumn )
			,JavaCast( "int", arguments.topRow )
			,JavaCast( "int", arguments.activePane )
		);
	}

	public void function autoSizeColumn( required workbook, required numeric column, boolean useMergedCells=false ){
		if( arguments.column LTE 0 )
			Throw( type=exceptionType, message="Invalid column value", detail="The value for column must be greater than or equal to 1." );
		/* Adjusts the width of the specified column to fit the contents. For performance reasons, this should normally be called only once per column. */
		var columnIndex = ( arguments.column -1 );
		if( isStreamingXmlFormat( arguments.workbook ) )
			getActiveSheet( arguments.workbook ).trackColumnForAutoSizing( JavaCast( "int", columnIndex ) );
		getActiveSheet( arguments.workbook ).autoSizeColumn( columnIndex, arguments.useMergedCells );
	}

	public void function cleanUpStreamingXml( required workbook ){
		if( isStreamingXmlFormat( arguments.workbook ) )
			arguments.workbook.dispose(); // SXSSF uses temporary files which MUST be cleaned up, see http://poi.apache.org/components/spreadsheet/how-to.html#sxssf
	}

	public void function clearCell( required workbook, required numeric row, required numeric column ){
		/* Clears the specified cell of all styles and values */
		var defaultStyle = arguments.workbook.getCellStyleAt( JavaCast( "short", 0 ) );
		var rowIndex = ( arguments.row -1 );
		var rowObject = getActiveSheet( arguments.workbook ).getRow( JavaCast( "int", rowIndex ) );
		if( IsNull( rowObject ) )
			return;
		var columnIndex = ( arguments.column -1 );
		var cell = rowObject.getCell( JavaCast( "int", columnIndex ) );
		if( IsNull( cell ) )
			return;
		cell.setCellStyle( defaultStyle );
		cell.setCellType( cell.CellType.BLANK );
	}

	public void function clearCellRange(
		required workbook
		,required numeric startRow
		,required numeric startColumn
		,required numeric endRow
		,required numeric endColumn
	){
		/* Clears the specified cell range of all styles and values */
		for( var rowNumber = arguments.startRow; rowNumber LTE arguments.endRow; rowNumber++ ){
			for( var columnNumber = arguments.startColumn; columnNumber LTE arguments.endColumn; columnNumber++ ){
				clearCell( arguments.workbook, rowNumber, columnNumber );
			}
		}
	}

	public void function createSheet( required workbook, string sheetName, overwrite=false ){
		if( arguments.KeyExists( "sheetName" ) )
			validateSheetName( arguments.sheetName );
		else
			arguments.sheetName = generateUniqueSheetName( arguments.workbook );
		if( !sheetExists( workbook=arguments.workbook, sheetName=arguments.sheetName ) ){
			arguments.workbook.createSheet( JavaCast( "String", arguments.sheetName ) );
			return;
		}
		/* sheet already exists with that name */
		if( !arguments.overwrite )
			Throw( type=exceptionType, message="Sheet name already exists", detail="A sheet with the name '#arguments.sheetName#' already exists in this workbook" );
		/* OK to replace the existing */
		var sheetIndexToReplace = arguments.workbook.getSheetIndex( JavaCast( "string", arguments.sheetName ) );
		deleteSheetAtIndex( arguments.workbook, sheetIndexToReplace );
		var newSheet = arguments.workbook.createSheet( JavaCast( "String", arguments.sheetName ) );
		var moveToIndex = sheetIndexToReplace;
		moveSheet( arguments.workbook, arguments.sheetName, moveToIndex );
	}

	public void function deleteColumn( required workbook,required numeric column ){
		if( arguments.column LTE 0 )
			Throw( type=exceptionType, message="Invalid column value", detail="The value for column must be greater than or equal to 1." );
			/* POI doesn't have remove column functionality, so iterate over all the rows and remove the column indicated */
		var rowIterator = getActiveSheet( arguments.workbook ).rowIterator();
		while( rowIterator.hasNext() ){
			var row = rowIterator.next();
			var cell = row.getCell( JavaCast( "int", ( arguments.column -1 ) ) );
			if( IsNull( cell ) )
				continue;
			row.removeCell( cell );
		}
	}

	public void function deleteColumns( required workbook, required string range ){
		/* Validate and extract the ranges. Range is a comma-delimited list of ranges, and each value can be either a single number or a range of numbers with a hyphen. */
		var allRanges = extractRanges( arguments.range );
		for( var thisRange in allRanges ){
			if( thisRange.startAt EQ thisRange.endAt ){
				/* Just one row */
				deleteColumn( arguments.workbook, thisRange.startAt );
				continue;
			}
			for( var columnNumber = thisRange.startAt; columnNumber LTE thisRange.endAt; columnNumber++ )
				deleteColumn( arguments.workbook, columnNumber );
		}
	}

	public void function deleteRow( required workbook, required numeric row ){
		/* Deletes the data from a row. Does not physically delete the row. */
		if( arguments.row LTE 0 )
			Throw( type=exceptionType, message="Invalid row value", detail="The value for row must be greater than or equal to 1." );
		var rowToDelete = ( arguments.row -1 );
		if( rowToDelete GTE getFirstRowNum( arguments.workbook ) AND rowToDelete LTE getLastRowNum( arguments.workbook ) ) //If this is a valid row, remove it
			getActiveSheet( arguments.workbook ).removeRow( getActiveSheet( arguments.workbook ).getRow( JavaCast( "int", rowToDelete ) ) );
	}

	public void function deleteRows( required workbook, required string range ){
		/* Validate and extract the ranges. Range is a comma-delimited list of ranges, and each value can be either a single number or a range of numbers with a hyphen. */
		var allRanges = extractRanges( arguments.range );
		for( var thisRange in allRanges ){
			if( thisRange.startAt EQ thisRange.endAt ){
				/* Just one row */
				deleteRow( arguments.workbook, thisRange.startAt );
				continue;
			}
			for( var rowNumber = thisRange.startAt; rowNumber LTE thisRange.endAt; rowNumber++ )
				deleteRow( arguments.workbook, rowNumber );
		}
	}

	public void function formatCell(
		required workbook
		,required struct format
		,required numeric row
		,required numeric column
		,any cellStyle
	){
		var cell = initializeCell( arguments.workbook, arguments.row, arguments.column );
		var style = arguments.cellStyle?: buildCellStyle( arguments.workbook, arguments.format );
		cell.setCellStyle( style );
	}

	public void function formatCellRange(
		required workbook
		,required struct format
		,required numeric startRow
		,required numeric endRow
		,required numeric startColumn
		,required numeric endColumn
		,any cellStyle
		){
		var style = arguments.cellStyle?: buildCellStyle( arguments.workbook, arguments.format );
		for( var rowNumber = arguments.startRow; rowNumber LTE arguments.endRow; rowNumber++ ){
			for( var columnNumber = arguments.startColumn; columnNumber LTE arguments.endColumn; columnNumber++ )
				formatCell( arguments.workbook, arguments.format, rowNumber, columnNumber, style );
		}
	}

	public void function formatColumn(
		required workbook
		,required struct format
		,required numeric column
		,any cellStyle
	){
		if( arguments.column LT 1 )
			Throw( type=exceptionType, message="Invalid column value", detail="The column value must be greater than 0" );
		var style = arguments.cellStyle?: buildCellStyle( arguments.workbook, arguments.format );
		var rowIterator = getActiveSheet( arguments.workbook ).rowIterator();
		var columnNumber = arguments.column;
		while( rowIterator.hasNext() ){
			var rowNumber = rowIterator.next().getRowNum() + 1;
			formatCell( arguments.workbook, arguments.format, rowNumber, columnNumber, style );
		}
	}

	public void function formatColumns(
		required workbook
		,required struct format
		,required string range
		,any cellStyle
	){
		/* Validate and extract the ranges. Range is a comma-delimited list of ranges, and each value can be either a single number or a range of numbers with a hyphen. */
		var allRanges = extractRanges( arguments.range );
		var style = arguments.cellStyle?: buildCellStyle( arguments.workbook, arguments.format );
		for( var thisRange in allRanges ){
			if( thisRange.startAt EQ thisRange.endAt ){
				/* Just one column */
				formatColumn( arguments.workbook, arguments.format, thisRange.startAt, style );
				continue;
			}
			for( var columnNumber = thisRange.startAt; columnNumber LTE thisRange.endAt; columnNumber++ )
				formatColumn( arguments.workbook, arguments.format, columnNumber, style );
		}
	}

	public void function formatRow(
		required workbook
		,required struct format
		,required numeric row
		,any cellStyle
	){
		var rowIndex = ( arguments.row -1 );
		var theRow = getActiveSheet( arguments.workbook ).getRow( rowIndex );
		if( IsNull( theRow ) )
			return;
		var style = arguments.cellStyle?: buildCellStyle( arguments.workbook, arguments.format );
		var cellIterator = theRow.cellIterator();
		while( cellIterator.hasNext() )
			formatCell( arguments.workbook, arguments.format, arguments.row, ( cellIterator.next().getColumnIndex() +1 ), style );
	}

	public void function formatRows(
		required workbook
		,required struct format
		,required string range
		,any cellStyle
	){
		/* Validate and extract the ranges. Range is a comma-delimited list of ranges, and each value can be either a single number or a range of numbers with a hyphen. */
		var allRanges = extractRanges( arguments.range );
		var style = arguments.cellStyle?: buildCellStyle( arguments.workbook, arguments.format );
		for( var thisRange in allRanges ){
			if( thisRange.startAt EQ thisRange.endAt ){
				/* Just one row */
				formatRow( arguments.workbook, arguments.format, thisRange.startAt, style );
				continue;
			}
			for( var rowNumber = thisRange.startAt; rowNumber LTE thisRange.endAt; rowNumber++ )
				formatRow( arguments.workbook, arguments.format, rowNumber, style );
		}
	}

	public any function getCellComment( required workbook, numeric row, numeric column ){
		if( arguments.KeyExists( "row" ) AND !arguments.KeyExists( "column" ) )
			Throw( type=exceptionType, message="Invalid argument combination", detail="If you specify the row you must also specify the column" );
		if( arguments.KeyExists( "column" ) AND !arguments.KeyExists( "row" ) )
			Throw( type=exceptionType, message="Invalid argument combination", detail="If you specify the column you must also specify the row" );
		if( arguments.KeyExists( "row" ) ){
			var cell = getCellAt( arguments.workbook, arguments.row, arguments.column );
			var commentObject = cell.getCellComment();
			if( !IsNull( commentObject ) ){
				return {
					author: commentObject.getAuthor()
					,comment: commentObject.getString().getString()
					,column: arguments.column
					,row: arguments.row
				};
			}
			return {};
		}
		/* TODO: Look into checking all sheets in the workbook */
		/* row and column weren't provided so loop over the whole sheet and return all the comments as an array of structs */
		var comments = [];
		var rowIterator = getActiveSheet( arguments.workbook ).rowIterator();
		while( rowIterator.hasNext() ){
			var cellIterator = rowIterator.next().cellIterator();
			while( cellIterator.hasNext() ){
				var commentObject = cellIterator.next().getCellComment();
				if( !IsNull( commentObject ) ){
					var comment = {
						author: commentObject.getAuthor()
						,comment: commentObject.getString().getString()
						,column: arguments.column
						,row: arguments.row
					};
					comments.Append( comment );
				}
			}
		}
		return comments;
	}

	public struct function getCellFormat( required workbook, required numeric row, required numeric column ){
		if( !cellExists( arguments.workbook, arguments.row, arguments.column ) )
			Throw( type=exceptionType, message="Invalid cell", detail="There doesn't appear to be a cell at row #row#, column #column#" );
		var cellStyle = getCellAt( arguments.workbook, arguments.row, arguments.column ).getCellStyle();
		var cellFont = arguments.workbook.getFontAt( cellStyle.getFontIndex() );
		if( isXmlFormat( arguments.workbook ) )
			var rgb = convertSignedRGBToPositiveTriplet( cellFont.getXSSFColor().getRGB() );
		else
			var rgb = IsNull( cellFont.getHSSFColor( arguments.workbook ) )? []: cellFont.getHSSFColor( arguments.workbook ).getTriplet();
		return {
			alignment: cellStyle.getAlignmentEnum().toString()
			,bold: cellFont.getBold()
			,bottomborder: cellStyle.getBorderBottomEnum().toString()
			,bottombordercolor: getRgbTripletForStyleColorFormat( arguments.workbook, cellStyle, "bottombordercolor" )
			,color: ArrayToList( rgb )
			,dataformat: cellStyle.getDataFormatString()
			,fgcolor: getRgbTripletForStyleColorFormat( arguments.workbook, cellStyle, "fgcolor" )
			,fillpattern: cellStyle.getFillPatternEnum().toString()
			,font: cellFont.getFontName()
			,fontsize: cellFont.getFontHeightInPoints()
			,indent: cellStyle.getIndention()
			,italic: cellFont.getItalic()
			,leftborder: cellStyle.getBorderLeftEnum().toString()
			,leftbordercolor: getRgbTripletForStyleColorFormat( arguments.workbook, cellStyle, "leftbordercolor" )
			,quoteprefixed: cellStyle.getQuotePrefixed()
			,rightborder: cellStyle.getBorderRightEnum().toString()
			,rightbordercolor: getRgbTripletForStyleColorFormat( arguments.workbook, cellStyle, "rightbordercolor" )
			,rotation: cellStyle.getRotation()
			,strikeout: cellFont.getStrikeout()
			,textwrap: cellStyle.getWrapText()
			,topborder: cellStyle.getBorderTopEnum().toString()
			,topbordercolor: getRgbTripletForStyleColorFormat( arguments.workbook, cellStyle, "topbordercolor" )
			,underline: getUnderlineFormatAsString( cellFont )
			,verticalalignment: cellStyle.getVerticalAlignmentEnum().toString()
		};
	}

	public any function getCellFormula( required workbook, numeric row, numeric column ){
		if( arguments.KeyExists( "row" ) AND arguments.KeyExists( "column" ) ){
			if( cellExists( arguments.workbook, arguments.row, arguments.column ) ){
				var cell = getCellAt( arguments.workbook, arguments.row, arguments.column );
				if( cellIsOfType( cell, "FORMULA" ) )
					return cell.getCellFormula();
				return "";
			}
		}
		//no row and column provided so return an array of structs containing formulas for the entire sheet
		var rowIterator = getActiveSheet( arguments.workbook ).rowIterator();
		var formulas = [];
		while( rowIterator.hasNext() ){
			var cellIterator = rowIterator.next().cellIterator();
			while( cellIterator.hasNext() ){
				var cell = cellIterator.next();
				var formulaStruct = {
					row: ( cell.getRowIndex() + 1 )
					,column: ( cell.getColumnIndex() + 1 )
				};
				try{
					formulaStruct.formula = cell.getCellFormula();
				}
				catch( any exception ){
					formulaStruct.formula = "";
				}
				if( formulaStruct.formula.Len() )
					formulas.Append( formulaStruct );
			}
		}
		return formulas;
	}

	public string function getCellType( required workbook, required numeric row, required numeric column ){
		if( !cellExists( arguments.workbook, arguments.row, arguments.column ) )
			return "";
		var rowIndex = ( arguments.row -1 );
		var columnIndex = ( arguments.column -1 );
		var rowObject = getActiveSheet( arguments.workbook ).getRow( JavaCast( "int", rowIndex ) );
		var cell = rowObject.getCell( JavaCast( "int", columnIndex ) );
		return cell.getCellTypeEnum().toString();
	}

	public any function getCellValue( required workbook, required numeric row, required numeric column ){
		if( !cellExists( arguments.workbook, arguments.row, arguments.column ) )
			return "";
		var rowIndex = ( arguments.row -1 );
		var columnIndex = ( arguments.column -1 );
		var rowObject = getActiveSheet( arguments.workbook ).getRow( JavaCast( "int", rowIndex ) );
		var cell = rowObject.getCell( JavaCast( "int", columnIndex ) );
		var formatter = getFormatter();
		if( cellIsOfType( cell, "FORMULA" ) ){
			var formulaEvaluator = arguments.workbook.getCreationHelper().createFormulaEvaluator();
			return formatter.formatCellValue( cell, formulaEvaluator );
		}
		return formatter.formatCellValue( cell );
	}

	public numeric function getColumnCount( required workbook, sheetNameOrNumber ){
		if( arguments.KeyExists( "sheetNameOrNumber" ) ){
			if( IsValid( "integer", arguments.sheetNameOrNumber ) AND IsNumeric( arguments.sheetNameOrNumber ) )
				setActiveSheetNumber( arguments.workbook, arguments.sheetNameOrNumber );
			else
				setActiveSheet( arguments.workbook, arguments.sheetNameOrNumber );
		}
		var sheet = getActiveSheet( arguments.workbook );
		var rowIterator = sheet.rowIterator();
		var result = 0;
		while( rowIterator.hasNext() ){
			var row = rowIterator.next();
			result = Max( result, row.getLastCellNum() );
		}
		return result;
	}

	public array function getPresetColorNames(){
		var presetEnum = loadClass( "org.apache.poi.hssf.util.HSSFColor$HSSFColorPredefined" );
		var result = [];
		for( var value in presetEnum.values() )
			result.Append( value.name() );
		result.Sort( "text" );//ACF2016 (not 2018) returns "YES" from a sort instead of the sorted array, so perform sort separately.
		return result;
	}

	public numeric function getRowCount( required workbook, sheetNameOrNumber ){
		if( arguments.KeyExists( "sheetNameOrNumber" ) ){
			if( IsValid( "integer", arguments.sheetNameOrNumber ) AND IsNumeric( arguments.sheetNameOrNumber ) )
				setActiveSheetNumber( arguments.workbook, arguments.sheetNameOrNumber );
			else
				setActiveSheet( arguments.workbook, arguments.sheetNameOrNumber );
		}
		var sheet = getActiveSheet( arguments.workbook );
		var lastRowIndex = getLastRowNum( arguments.workbook, sheet );
		if( lastRowIndex == -1 )// empty
			return 0;
		return lastRowIndex +1;
	}

	public void function hideColumn( required workbook, required numeric column ){
		toggleColumnHidden( arguments.workbook, arguments.column, true );
	}

	public void function hideRow( required workbook, required numeric row ){
		toggleRowHidden( arguments.workbook, arguments.row, true );
	}

	public struct function info( required workbookOrPath ){
		/*
		properties returned in the struct are:
			* AUTHOR
			* CATEGORY
			* COMMENTS
			* CREATIONDATE
			* LASTEDITED
			* LASTAUTHOR
			* LASTSAVED
			* KEYWORDS
			* MANAGER
			* COMPANY
			* SUBJECT
			* TITLE
			* SHEETS
			* SHEETNAMES
			* SPREADSHEETTYPE
		*/
		if( this.isSpreadsheetObject( arguments[ 1 ] ) ) //use this scope to avoid clash with ACF built-in function
			var workbook = arguments[ 1 ];
		else
			var workbook = workbookFromFile( arguments[ 1 ] );
		//format specific metadata
		var info = isBinaryFormat( workbook )? binaryInfo( workbook ): xmlInfo( workbook );
		//common properties
		info.sheets = workbook.getNumberOfSheets();
		var sheetnames = [];
		if( IsNumeric( info.sheets ) ){
			for( var i = 1; i LTE info.sheets; i++ )
				sheetnames.Append( workbook.getSheetName( JavaCast( "int", ( i -1 ) ) ) );
			info.sheetnames = sheetnames.ToList();
		}
		info.spreadSheetType = isXmlFormat( workbook )? "Excel (2007)": "Excel";
		return info;
	}

	public boolean function isBinaryFormat( required workbook ){
		return arguments.workbook.getClass().getCanonicalName() IS variables.HSSFWorkbookClassName;
	}

	public boolean function isColumnHidden( required workbook, required numeric column ){
		return getActiveSheet( arguments.workbook ).isColumnHidden( arguments.column - 1 );
	}

	public boolean function isRowHidden( required workbook, required numeric row ){
		var rowIndex = ( arguments.row - 1 );
		return getActiveSheet( arguments.workbook ).getRow( JavaCast( "int", rowIndex ) ).getZeroHeight();
	}

	public boolean function isSpreadsheetFile( required string path ){
		if( !FileExists( arguments.path ) )
			Throw( type=exceptionType, message="Non-existent file", detail="Cannot find the file #arguments.path#." );
		try{
			var workbook = workbookFromFile( arguments.path );
		}
		catch( cfsimplicity.lucee.spreadsheet.invalidFile exception ){
			return false;
		}
		return true;
	}

	public boolean function isSpreadsheetObject( required object ){
		return isBinaryFormat( arguments.object ) OR isXmlFormat( arguments.object );
	}

	public boolean function isXmlFormat( required workbook ){
		//CF2016 doesn't support [].Find( needle );
		return ArrayFind( [ variables.XSSFWorkbookClassName, variables.SXSSFWorkbookClassName ], arguments.workbook.getClass().getCanonicalName() );
	}

	public boolean function isStreamingXmlFormat( required workbook ){
		return arguments.workbook.getClass().getCanonicalName() IS variables.SXSSFWorkbookClassName;
	}

	public void function mergeCells(
		required workbook
		,required numeric startRow
		,required numeric endRow
		,required numeric startColumn
		,required numeric endColumn
		,boolean emptyInvisibleCells=false
	){
		if( arguments.startRow LT 1 OR arguments.startRow GT arguments.endRow )
			Throw( type=exceptionType, message="Invalid startRow or endRow", detail="Row values must be greater than 0 and the startRow cannot be greater than the endRow." );
		if( arguments.startColumn LT 1 OR arguments.startColumn GT arguments.endColumn )
			Throw( type=exceptionType, message="Invalid startColumn or endColumn", detail="Column values must be greater than 0 and the startColumn cannot be greater than the endColumn." );
		var cellRangeAddress = loadClass( "org.apache.poi.ss.util.CellRangeAddress" ).init(
			JavaCast( "int", ( arguments.startRow - 1 ) )
			,JavaCast( "int", ( arguments.endRow - 1 ) )
			,JavaCast( "int", ( arguments.startColumn - 1 ) )
			,JavaCast( "int", ( arguments.endColumn - 1 ) )
		);
		getActiveSheet( arguments.workbook ).addMergedRegion( cellRangeAddress );
		if( !arguments.emptyInvisibleCells )
			return;
		// stash the value to retain
		var visibleValue = getCellValue( arguments.workbook, arguments.startRow, arguments.startColumn );
		//empty all cells in the merged region
		setCellRangeValue( arguments.workbook, "", arguments.startRow, arguments.endRow, arguments.startColumn, arguments.endColumn );
		//restore the stashed value
		setCellValue( arguments.workbook, visibleValue, arguments.startRow, arguments.startColumn );
	}

	public any function new(
		string sheetName="Sheet1"
		,boolean xmlFormat=false
		,boolean streamingXml=false
		,numeric streamingWindowSize
	){
		if( arguments.streamingXml && !arguments.xmlFormat )
			arguments.xmlFormat = true;
		var workbook = createWorkBook( argumentCollection=arguments );
		createSheet( workbook, arguments.sheetName, arguments.xmlFormat );
		setActiveSheet( workbook, arguments.sheetName );
		return workbook;
	}

	public any function newXls( string sheetName="Sheet1" ){
		return new( sheetName=arguments.sheetName, xmlFormat=false );
	}

	public any function newXlsx( string sheetName="Sheet1" ){
		return new( sheetName=arguments.sheetName, xmlFormat=true );
	}

	public any function newStreamingXlsx( string sheetName="Sheet1", numeric streamingWindowSize=100 ){
		return new( sheetName=arguments.sheetName, xmlFormat=true, streamingXml=true, streamingWindowSize=arguments.streamingWindowSize );
	}

	public any function read(
		required string src
		,string format
		,string columns
		,string columnNames
		,numeric headerRow
		,string rows
		,string sheetName
		,numeric sheetNumber // 1-based
		,boolean includeHeaderRow=false
		,boolean includeBlankRows=false
		,boolean fillMergedCellsWithVisibleValue=false
		,boolean includeHiddenColumns=true
		,boolean includeRichTextFormatting=false
		,string password
	){
		if( arguments.KeyExists( "query" ) )
			Throw( type=exceptionType, message="Invalid argument 'query'.", detail="Just use format='query' to return a query object" );
		if( arguments.KeyExists( "format" ) AND !ListFindNoCase( "query,html,csv", arguments.format ) )
			Throw( type=exceptionType, message="Invalid format", detail="Supported formats are: 'query', 'html' and 'csv'" );
		if( arguments.KeyExists( "sheetName" ) AND arguments.KeyExists( "sheetNumber" ) )
			Throw( type=exceptionType, message="Cannot provide both sheetNumber and sheetName arguments", detail="Only one of either 'sheetNumber' or 'sheetName' arguments may be provided." );
		if( !FileExists( arguments.src ) )
			Throw( type=exceptionType, message="Non-existent file", detail="Cannot find the file #arguments.src#." );
		var passwordProtected = ( arguments.KeyExists( "password") AND !password.Trim().IsEmpty() );
		var workbook = passwordProtected? workbookFromFile( arguments.src, password ): workbookFromFile( arguments.src );
		if( arguments.KeyExists( "sheetName" ) )
			setActiveSheet( workbook=workbook, sheetName=arguments.sheetName );
		if( !arguments.KeyExists( "format" ) )
			return workbook;
		var args = {
			workbook: workbook
		};
		if( arguments.KeyExists( "sheetName" ) )
			args.sheetName = arguments.sheetName;
		if( arguments.KeyExists( "sheetNumber" ) )
			args.sheetNumber = arguments.sheetNumber;
		if( arguments.KeyExists( "headerRow" ) ){
			args.headerRow = arguments.headerRow;
			args.includeHeaderRow = arguments.includeHeaderRow;
		}
		if( arguments.KeyExists( "rows" ) )
			args.rows = arguments.rows;
		if( arguments.KeyExists( "columns" ) )
			args.columns = arguments.columns;
		if( arguments.KeyExists( "columnNames" ) )
			args.columnNames = arguments.columnNames;
		args.includeBlankRows = arguments.includeBlankRows;
		args.fillMergedCellsWithVisibleValue = arguments.fillMergedCellsWithVisibleValue;
		args.includeHiddenColumns = arguments.includeHiddenColumns;
		args.includeRichTextFormatting = arguments.includeRichTextFormatting;
		var generatedQuery = sheetToQuery( argumentCollection=args );
		if( arguments.format IS "query" )
			return generatedQuery;
		var args = {
			query: generatedQuery
		};
		if( arguments.KeyExists( "headerRow" ) ){
			args.headerRow = arguments.headerRow;
			args.includeHeaderRow = arguments.includeHeaderRow;
		}
		switch( arguments.format ){
			case "csv": return queryToCsv( argumentCollection=args );
			case "html": return queryToHtml( argumentCollection=args );
		}
	}

	public binary function readBinary( required workbook ){
		var baos = CreateObject( "Java", "org.apache.commons.io.output.ByteArrayOutputStream" ).init();
		arguments.workbook.write( baos );
		baos.flush();
		return baos.toByteArray();
	}

	public void function removePrintGridlines( required workbook ){
		getActiveSheet( arguments.workbook ).setPrintGridlines( JavaCast( "boolean", false ) );
	}

	public void function removeSheet( required workbook, required string sheetName ){
		validateSheetName( arguments.sheetName );
		validateSheetExistsWithName( arguments.workbook, arguments.sheetName );
		arguments.sheetNumber = ( arguments.workbook.getSheetIndex( arguments.sheetName ) +1 );
		var sheetIndex = ( sheetNumber -1 );
		deleteSheetAtIndex( arguments.workbook, sheetIndex );
	}

	public void function removeSheetNumber( required workbook, required numeric sheetNumber ){
		validateSheetNumber( arguments.workbook, arguments.sheetNumber );
		var sheetIndex = ( arguments.sheetNumber -1 );
		deleteSheetAtIndex( arguments.workbook, sheetIndex );
	}

	public void function renameSheet( required workbook, required string sheetName, required numeric sheetNumber ){
		validateSheetName( arguments.sheetName );
		validateSheetNumber( arguments.workbook, arguments.sheetNumber );
		var sheetIndex = ( arguments.sheetNumber -1 );
		var foundAt = arguments.workbook.getSheetIndex( JavaCast( "string", arguments.sheetName ) );
		if( ( foundAt GT 0 ) AND ( foundAt NEQ sheetIndex ) )
			Throw( type=exceptionType, message="Invalid Sheet Name [#arguments.sheetName#]", detail="The workbook already contains a sheet named [#sheetName#]. Sheet names must be unique" );
		arguments.workbook.setSheetName( JavaCast( "int", sheetIndex ), JavaCast( "string", arguments.sheetName ) );
	}

	public void function setActiveSheet( required workbook, string sheetName, numeric sheetNumber ){
		validateSheetNameOrNumberWasProvided( argumentCollection=arguments );
		if( arguments.KeyExists( "sheetName" ) ){
			validateSheetExistsWithName( arguments.workbook, arguments.sheetName );
			arguments.sheetNumber = ( arguments.workbook.getSheetIndex( JavaCast( "string", arguments.sheetName ) ) + 1 );
		}
		validateSheetNumber( arguments.workbook, arguments.sheetNumber );
		arguments.workbook.setActiveSheet( JavaCast( "int", ( arguments.sheetNumber - 1 ) ) );
	}

	public void function setActiveSheetNumber( required workbook, numeric sheetNumber ){
		setActiveSheet( workbook=arguments.workbook, sheetNumber=arguments.sheetNumber );
	}

	public void function setCellComment(
		required workbook
		,required struct comment
		,required numeric row
		,required numeric column
	){
		/*
		The comment struct may contain the following keys:
			* anchor
			* author
			* bold
			* color
			* comment
			* fillcolor
			* font
			* horizontalalignment
			* italic
			* linestyle
			* linestylecolor
			* size
			* strikeout
			* underline
			* verticalalignment
			* visible
		 */
		var drawingPatriarch = getActiveSheet( arguments.workbook ).createDrawingPatriarch();
		var commentString = loadClass( "org.apache.poi.hssf.usermodel.HSSFRichTextString" ).init( JavaCast( "string", arguments.comment.comment ) );
		var javaColorRGB = 0;
		if( arguments.comment.KeyExists( "anchor" ) )
			var clientAnchor = loadClass( "org.apache.poi.hssf.usermodel.HSSFClientAnchor" ).init(
				JavaCast( "int", 0 )
				,JavaCast( "int", 0 )
				,JavaCast( "int", 0 )
				,JavaCast( "int", 0 )
				,JavaCast( "short", ListGetAt( arguments.comment.anchor, 1 ) )
				,JavaCast( "int", ListGetAt( arguments.comment.anchor, 2 ) )
				,JavaCast( "short", ListGetAt( arguments.comment.anchor, 3 ) )
				,JavaCast( "int", ListGetAt( arguments.comment.anchor, 4 ) )
			);
		else
			var clientAnchor = loadClass( "org.apache.poi.hssf.usermodel.HSSFClientAnchor" ).init(
				JavaCast( "int", 0 )
				,JavaCast( "int", 0 )
				,JavaCast( "int", 0 )
				,JavaCast( "int", 0 )
				,JavaCast( "short", arguments.column )
				,JavaCast( "int", arguments.row )
				,JavaCast( "short", ( arguments.column +2 ) )
				,JavaCast( "int", ( arguments.row +2 ) )
			);
		var commentObject = drawingPatriarch.createComment( clientAnchor );
		if( arguments.comment.KeyExists( "author" ) )
			commentObject.setAuthor( JavaCast( "string", arguments.comment.author ) );
		/* If we're going to do anything font related, need to create a font. Didn't really want to create it above since it might not be needed.  */
		if( arguments.comment.KeyExists( "bold" )
				OR arguments.comment.KeyExists( "color" )
				OR arguments.comment.KeyExists( "font" )
				OR arguments.comment.KeyExists( "italic" )
				OR arguments.comment.KeyExists( "size" )
				OR arguments.comment.KeyExists( "strikeout" )
				OR arguments.comment.KeyExists( "underline" )
		){
			var font = workbook.createFont();
			if( arguments.comment.KeyExists( "bold" ) )
				font.setBold( JavaCast( "boolean", arguments.comment.bold ) );
			if( arguments.comment.KeyExists( "color" ) )
				font.setColor( getColor( workbook, arguments.comment.color ) );
			if( arguments.comment.KeyExists( "font" ) )
				font.setFontName( JavaCast( "string", arguments.comment.font ) );
			if( arguments.comment.KeyExists( "italic" ) )
				font.setItalic( JavaCast( "string", arguments.comment.italic ) );
			if( arguments.comment.KeyExists( "size" ) )
				font.setFontHeightInPoints( JavaCast( "int", arguments.comment.size ) );
			if( arguments.comment.KeyExists( "strikeout" ) )
				font.setStrikeout( JavaCast( "boolean", arguments.comment.strikeout ) );
			if( arguments.comment.KeyExists( "underline" ) )
				font.setUnderline( JavaCast( "boolean", arguments.comment.underline ) );
			arguments.commentString.applyFont( font );
		}
		if( arguments.comment.KeyExists( "fillColor" ) ){
			javaColorRGB = getJavaColorRGB( arguments.comment.fillColor );
			commentObject.setFillColor(
				JavaCast( "int", javaColorRGB.red )
				,JavaCast( "int", javaColorRGB.green )
				,JavaCast( "int", javaColorRGB.blue )
			);
		}
		/*
			Horizontal alignment can be left, center, right, justify, or distributed. Note that the constants on the Java class are slightly different in some cases:
			'center'=CENTERED
			'justify'=JUSTIFIED
		 */
		if( arguments.comment.KeyExists( "horizontalAlignment" ) ){
			if( arguments.comment.horizontalAlignment.UCase() IS "CENTER" )
				arguments.comment.horizontalAlignment="CENTERED";
			if( arguments.comment.horizontalAlignment.UCase() IS "JUSTIFY" )
				arguments.comment.horizontalAlignment="JUSTIFIED";
			commentObject.setHorizontalAlignment( JavaCast( "int", commentObject[ "HORIZONTAL_ALIGNMENT_" & arguments.comment.horizontalalignment.UCase() ] ) );
		}
		/*
		Valid values for linestyle are:
				* solid
				* dashsys
				* dashdotsys
				* dashdotdotsys
				* dotgel
				* dashgel
				* longdashgel
				* dashdotgel
				* longdashdotgel
				* longdashdotdotgel
		 */
		if( arguments.comment.KeyExists( "lineStyle" ) )
		 	commentObject.setLineStyle( JavaCast( "int", commentObject[ "LINESTYLE_" & arguments.comment.lineStyle.UCase() ] ) );
		if( arguments.comment.KeyExists( "lineStyleColor" ) ){
			javaColorRGB = getJavaColorRGB( arguments.comment.lineStyleColor );
			commentObject.setLineStyleColor(
				JavaCast( "int", javaColorRGB.red )
				,JavaCast( "int", javaColorRGB.green )
				,JavaCast( "int", javaColorRGB.blue )
			);
		}
		/* Vertical alignment can be top, center, bottom, justify, and distributed. Note that center and justify are DIFFERENT than the constants for horizontal alignment, which are CENTERED and JUSTIFIED. */
		if( arguments.comment.KeyExists( "verticalAlignment" ) )
			commentObject.setVerticalAlignment( JavaCast( "int", commentObject[ "VERTICAL_ALIGNMENT_" & arguments.comment.verticalAlignment.UCase() ] ) );
		if( arguments.comment.KeyExists( "visible" ) )
			commentObject.setVisible( JavaCast( "boolean", arguments.comment.visible ) );//doesn't seem to work
		commentObject.setString( commentString );
		var cell = initializeCell( arguments.workbook, arguments.row, arguments.column );
		cell.setCellComment( commentObject );
	}

	public void function setCellFormula(
		required workbook
		,required string formula
		,required numeric row
		,required numeric column
	){
		//Automatically create the cell if it does not exist, instead of throwing an error
		var cell = initializeCell( arguments.workbook, arguments.row, arguments.column );
		cell.setCellFormula( JavaCast( "string", arguments.formula ) );
	}

	public void function setCellValue( required workbook, required value, required numeric row, required numeric column, string type ){
		//Automatically create the cell if it does not exist, instead of throwing an error
		var args = {
			workbook: arguments.workbook
			,cell: initializeCell( arguments.workbook, arguments.row, arguments.column )
			,value: arguments.value
		};
		if( arguments.KeyExists( "type" ) )
			args.type = arguments.type;
		setCellValueAsType( argumentCollection=args );
	}

	public void function setCellRangeValue(
		required workbook
		,required value
		,required numeric startRow
		,required numeric endRow
		,required numeric startColumn
		,required numeric endColumn
	){
		/* Sets the same value to a range of cells */
		for( var rowNumber = arguments.startRow; rowNumber LTE arguments.endRow; rowNumber++ ){
			for( var columnNumber = arguments.startColumn; columnNumber LTE arguments.endColumn; columnNumber++ )
				setCellValue( arguments.workbook, arguments.value, rowNumber, columnNumber );
		}
	}

	public void function setColumnWidth( required workbook, required numeric column, required numeric width ){
		var columnIndex = ( arguments.column -1 );
		getActiveSheet( arguments.workbook ).setColumnWidth( JavaCast( "int", columnIndex ), JavaCast( "int", ( arguments.width * 256 ) ) );
	}

	public void function setFitToPage( required workbook, required boolean state, numeric pagesWide, numeric pagesHigh ){
		var sheet = getActiveSheet( arguments.workbook );
		sheet.setFitToPage( JavaCast( "boolean", arguments.state ) );
		sheet.setAutoBreaks( JavaCast( "boolean", arguments.state ) ); //seems dependent on this matching
		if( !arguments.state )
			return;
		if( arguments.KeyExists( "pagesWide" ) && IsValid( "integer", arguments.pagesWide ) )
			sheet.getPrintSetup().setFitWidth( JavaCast( "short", arguments.pagesWide ) );
		if( arguments.KeyExists( "pagesWide" ) && IsValid( "integer", arguments.pagesHigh ) )
			sheet.getPrintSetup().setFitHeight( JavaCast( "short", arguments.pagesHigh ) );
	}

	public void function setFooter(
		required workbook
		,string leftFooter=""
		,string centerFooter=""
		,string rightFooter=""
	){
		var footer = getActiveSheet( arguments.workbook ).getFooter();
		if( !arguments.centerFooter.IsEmpty() )
			footer.setCenter( JavaCast( "string", arguments.centerFooter ) );
		if( !arguments.leftFooter.IsEmpty() )
			footer.setleft( JavaCast( "string", arguments.leftFooter ) );
		if( !arguments.rightFooter.IsEmpty() )
			footer.setright( JavaCast( "string", arguments.rightFooter ) );
	}

	public void function setHeader(
		required workbook
		,string leftHeader=""
		,string centerHeader=""
		,string rightHeader=""
	){
		var header = getActiveSheet( arguments.workbook ).getHeader();
		if( !arguments.centerHeader.IsEmpty() )
			header.setCenter( JavaCast( "string", arguments.centerHeader ) );
		if( !arguments.leftHeader.IsEmpty() )
			header.setleft( JavaCast( "string", arguments.leftHeader ) );
		if( !arguments.rightHeader.IsEmpty() )
			header.setright( JavaCast( "string", arguments.rightHeader ) );
	}

	public void function setReadOnly( required workbook, required string password ){
		if( isXmlFormat( arguments.workbook ) )
			Throw( type=exceptionType, message="setReadOnly not supported for XML workbooks", detail="The setReadOnly() method only works on binary 'xls' workbooks." );
		// writeProtectWorkbook takes both a user name and a password, just making up a user name
		arguments.workbook.writeProtectWorkbook( JavaCast( "string", arguments.password ), JavaCast( "string", "user" ) );
	}

	public void function setRepeatingColumns( required workbook, required string columnRange ){
		arguments.columnRange = arguments.columnRange.Trim();
		if( !IsValid( "regex", arguments.columnRange,"[A-Za-z]:[A-Za-z]" ) )
			Throw( type=exceptionType, message="Invalid columnRange argument", detail="The 'columnRange' argument should be in the form 'A:B'" );
		var cellRangeAddress = getCellRangeAddressFromReference( arguments.columnRange );
		getActiveSheet( arguments.workbook ).setRepeatingColumns( cellRangeAddress );
	}

	public void function setRepeatingRows( required workbook, required string rowRange ){
		arguments.rowRange = arguments.rowRange.Trim();
		if( !IsValid( "regex", arguments.rowRange,"\d+:\d+" ) )
			Throw( type=exceptionType, message="Invalid rowRange argument", detail="The 'rowRange' argument should be in the form 'n:n', e.g. '1:5'" );
		var cellRangeAddress = getCellRangeAddressFromReference( arguments.rowRange );
		getActiveSheet( arguments.workbook ).setRepeatingRows( cellRangeAddress );
	}

	public void function setRowHeight( required workbook, required numeric row, required numeric height ){
		var rowIndex = ( arguments.row -1 );
		getActiveSheet( arguments.workbook ).getRow( JavaCast( "int", rowIndex ) ).setHeightInPoints( JavaCast( "int", arguments.height ) );
	}

	public void function setSheetTopMargin( required workbook, required numeric marginSize, string sheetName, numeric sheetNumber ){
		var sheet = getSheetByNameOrNumber( argumentCollection=arguments );
		sheet.setMargin( sheet.TopMargin, arguments.marginSize );
	}

	public void function setSheetBottomMargin( required workbook, required numeric marginSize, string sheetName, numeric sheetNumber ){
		var sheet = getSheetByNameOrNumber( argumentCollection=arguments );
		sheet.setMargin( sheet.BottomMargin, arguments.marginSize );
	}

	public void function setSheetLeftMargin( required workbook, required numeric marginSize, string sheetName, numeric sheetNumber ){
		var sheet = getSheetByNameOrNumber( argumentCollection=arguments );
		sheet.setMargin( sheet.LeftMargin, arguments.marginSize );
	}

	public void function setSheetRightMargin( required workbook, required numeric marginSize, string sheetName, numeric sheetNumber ){
		var sheet = getSheetByNameOrNumber( argumentCollection=arguments );
		sheet.setMargin( sheet.RightMargin, arguments.marginSize );
	}

	public void function setSheetHeaderMargin( required workbook, required numeric marginSize, string sheetName, numeric sheetNumber ){
		var sheet = getSheetByNameOrNumber( argumentCollection=arguments );
		sheet.setMargin( sheet.HeaderMargin, arguments.marginSize );
	}

	public void function setSheetFooterMargin( required workbook, required numeric marginSize, string sheetName, numeric sheetNumber ){
		var sheet = getSheetByNameOrNumber( argumentCollection=arguments );
		sheet.setMargin( sheet.FooterMargin, arguments.marginSize );
	}

	public void function setSheetPrintOrientation( required workbook, required string mode, string sheetName, numeric sheetNumber ){
		if( !ListFindNoCase( "landscape,portrait", arguments.mode ) )
			Throw( type=exceptionType, message="Invalid mode argument", detail="#mode# is not a valid 'mode' argument. Use 'portrait' or 'landscape'" );
		var setToLandscape = ( LCase( arguments.mode ) IS "landscape" );
		var sheet = getSheetByNameOrNumber( argumentCollection=arguments );
		sheet.getPrintSetup().setLandscape( JavaCast( "boolean", setToLandscape ) );
	}

	public void function shiftColumns( required workbook, required numeric start, numeric end=arguments.start, numeric offset=1 ){
		if( arguments.start LTE 0 )
			Throw( type=exceptionType, message="Invalid start value", detail="The start value must be greater than or equal to 1" );
		if( arguments.KeyExists( "end" ) AND ( ( arguments.end LTE 0 ) OR ( arguments.end LT arguments.start ) ) )
			Throw( type=exceptionType, message="Invalid end value", detail="The end value must be greater than or equal to the start value" );
		var rowIterator = getActiveSheet( arguments.workbook ).rowIterator();
		var startIndex = ( arguments.start -1 );
		var endIndex = arguments.KeyExists( "end" )? ( arguments.end -1 ): startIndex;
		while( rowIterator.hasNext() ){
			var row = rowIterator.next();
			if( arguments.offset GT 0 ){
				for( var i = endIndex; i GTE startIndex; i-- ){
					var tempCell = row.getCell( JavaCast( "int", i ) );
					var cell = createCell( row, i + arguments.offset );
					if( !IsNull( tempCell ) ){
						setCellValueAsType( arguments.workbook, cell, getCellValueAsType( arguments.workbook, tempCell ) );
						cell.setCellStyle( tempCell.getCellStyle() );
					}
				}
			}
			else {
				for( var i = startIndex; i LTE endIndex; i++ ){
					var tempCell = row.getCell( JavaCast( "int", i ) );
					var cell = createCell( row, i + arguments.offset );
					if( !IsNull( tempCell ) ){
						setCellValueAsType( workbook, cell, getCellValueAsType( workbook, tempCell ) );
						cell.setCellStyle( tempCell.getCellStyle() );
					}
				}
			}
		}
		// clean up any columns that need to be deleted after the shift
		var numberColsShifted = ( ( endIndex-startIndex ) +1 );
		var numberColsToDelete = Abs( arguments.offset );
		if( numberColsToDelete GT numberColsShifted )
			numberColsToDelete = numberColsShifted;
		if( arguments.offset GT 0 ){
			var stopValue = ( ( startIndex + numberColsToDelete ) -1 );
			for( var i = startIndex; i LTE stopValue; i++ )
				deleteColumn( workbook, ( i +1 ) );
			return;
		}
		var stopValue = ( ( endIndex - numberColsToDelete ) +1 );
		for( var i = endIndex; i GTE stopValue; i-- )
			deleteColumn( workbook, ( i +1 ) );
	}

	public void function shiftRows( required workbook, required numeric start, numeric end=arguments.start, numeric offset=1 ){
		getActiveSheet( arguments.workbook ).shiftRows(
			JavaCast( "int", ( arguments.start - 1 ) )
			,JavaCast( "int", ( arguments.end - 1 ) )
			,JavaCast( "int", arguments.offset )
		);
	}

	public void function showColumn( required workbook, required numeric column ){
		toggleColumnHidden( arguments.workbook, arguments.column, false );
	}

	public void function showRow( required workbook, required numeric row ){
		toggleRowHidden( arguments.workbook, arguments.row, false );
	}

	public void function write(
		required workbook
		,required string filepath
		,boolean overwrite=false
		,string password
		,string algorithm="agile"
	){
		if( !arguments.overwrite AND FileExists( arguments.filepath ) )
			Throw( type=exceptionType, message="File already exists", detail="The file path specified already exists. Use 'overwrite=true' if you wish to overwrite it." );
		var passwordProtect = ( arguments.KeyExists( "password" ) AND !arguments.password.Trim().IsEmpty() );
		if( passwordProtect AND !engineSupportsWriteEncryption )
			Throw( type=exceptionType, message="Password protection is not supported for Adobe ColdFusion", detail="Password protection currently only works in Lucee, not ColdFusion" );
		if( passwordProtect AND isBinaryFormat( arguments.workbook ) )
			Throw( type=exceptionType, message="Whole file password protection is not supported for binary workbooks", detail="Password protection only works with XML ('xlsx') workbooks." );
		try{
			lock name="#arguments.filepath#" timeout=5{
				var outputStream = CreateObject( "java", "java.io.FileOutputStream" ).init( arguments.filepath );
				arguments.workbook.write( outputStream );
				outputStream.flush();
			}
		}
		finally{
			// always close the stream. otherwise file may be left in a locked state if an unexpected error occurs
			if( local.KeyExists( "outputStream" ) )
				outputStream.close();
			cleanUpStreamingXml( arguments.workbook );
		}
		if( passwordProtect )
			encryptFile( arguments.filepath, arguments.password, arguments.algorithm );
	}

	/* END PUBLIC API */

	/* PRIVATE METHODS */

	private void function addInfoBinary( required workbook, required struct info ){
		arguments.workbook.createInformationProperties(); // creates the following if missing
		var documentSummaryInfo = arguments.workbook.getDocumentSummaryInformation();
		var summaryInfo = arguments.workbook.getSummaryInformation();
		for( var key in arguments.info ){
			var value = JavaCast( "string", arguments.info[ key ] );
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

	private void function addInfoXml( required workbook, required struct info ){
		var documentProperties = arguments.workbook.getProperties().getExtendedProperties().getUnderlyingProperties();
		var coreProperties = arguments.workbook.getProperties().getCoreProperties();
		for( var key in arguments.info ){
			var value = JavaCast( "string", arguments.info[ key ] );
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

	private void function addRowToSheetData(
		required workbook
		,required struct sheet
		,required numeric rowIndex
		,boolean includeRichTextFormatting=false
	){
		if( ( arguments.rowIndex EQ arguments.sheet.headerRowIndex ) AND !arguments.sheet.includeHeaderRow )
			return;
		var rowData = [];
		var row = arguments.sheet.object.getRow( JavaCast( "int", arguments.rowIndex ) );
		if( IsNull( row ) ){
			if( arguments.sheet.includeBlankRows )
				arguments.sheet.data.Append( rowData );
			return;
		}
		if( rowIsEmpty( row ) AND !arguments.sheet.includeBlankRows )
			return;
		rowData = getRowData( arguments.workbook, row, arguments.sheet.columnRanges, arguments.includeRichTextFormatting );
		arguments.sheet.data.Append( rowData );
		if( !arguments.sheet.columnRanges.Len() ){
			var rowColumnCount = row.GetLastCellNum();
			arguments.sheet.totalColumnCount = Max( arguments.sheet.totalColumnCount, rowColumnCount );
		}
	}

	private struct function binaryInfo( required workbook ){
		var documentProperties = arguments.workbook.getDocumentSummaryInformation();
		var coreProperties = arguments.workbook.getSummaryInformation();
		return {
			author: coreProperties.getAuthor()?:""
			,category: documentProperties.getCategory()?:""
			,comments: coreProperties.getComments()?:""
			,creationDate: coreProperties.getCreateDateTime()?:""
			,lastEdited: ( coreProperties.getEditTime() EQ 0 )? "": CreateObject( "java", "java.util.Date" ).init( coreProperties.getEditTime() )
			,subject: coreProperties.getSubject()?:""
			,title: coreProperties.getTitle()?:""
			,lastAuthor: coreProperties.getLastAuthor()?:""
			,keywords: coreProperties.getKeywords()?:""
			,lastSaved: coreProperties.getLastSaveDateTime()?:""
			,manager: documentProperties.getManager()?:""
			,company: documentProperties.getCompany()?:""
		};
	}

	private boolean function cellExists( required workbook, required numeric rowNumber, required numeric columnNumber ){
		var rowIndex = ( arguments.rowNumber -1 );
		var columnIndex = ( arguments.columnNumber -1 );
		var checkRow = getActiveSheet( arguments.workbook ).getRow( JavaCast( "int", rowIndex ) );
		return !IsNull( checkRow ) AND !IsNull( checkRow.getCell( JavaCast( "int", columnIndex ) ) );
	}

	private boolean function cellIsOfType( required cell, required string type ){
		var cellType = arguments.cell.getCellType();
		return ObjectEquals( cellType, cellType[ arguments.type ] );
	}

	private numeric function columnCountFromRanges( required array ranges ){
		var result = 0;
		for( var thisRange in arguments.ranges ){
			for( var i = thisRange.startAt; i LTE thisRange.endAt; i++ )
				result++;
		}
		return result;
	}

	private array function convertSignedRGBToPositiveTriplet( required any signedRGB ){
		// When signed, values of 128+ are negative: convert then to positive values
		var result = [];
		for( var i=1; i LTE 3; i++ ){
			result.Append( ( arguments.signedRGB[ i ] < 0 )? ( arguments.signedRGB[ i ] + 256 ): arguments.signedRGB[ i ] );
		}
		return result;
	}

	private any function createCell( required row, numeric cellNum=arguments.row.getLastCellNum(), overwrite=true ){
		/* get existing cell (if any)  */
		var cell = arguments.row.getCell( JavaCast( "int", arguments.cellNum ) );
		if( arguments.overwrite AND !IsNull( cell ) )
			arguments.row.removeCell( cell );/* forcibly remove the existing cell  */
		if( arguments.overwrite OR IsNull( cell ) )
			cell = arguments.row.createCell( JavaCast( "int", arguments.cellNum ) );/* create a brand new cell  */
		return cell;
	}

	private any function createRow( required workbook, numeric rowNum=getNextEmptyRow( arguments.workbook ), boolean overwrite=true ){
		/* get existing row (if any)  */
		var sheet = getActiveSheet( arguments.workbook );
		var row = sheet.getRow( JavaCast( "int", rowNum ) );
		if( arguments.overwrite AND !IsNull( row ) )
			sheet.removeRow( row ); /* forcibly remove existing row and all cells  */
		if( arguments.overwrite OR IsNull( sheet.getRow( JavaCast( "int", rowNum ) ) ) ){
			try{
				row = sheet.createRow( JavaCast( "int", rowNum ) );
			}
			catch( java.lang.IllegalArgumentException exception ){
				if( exception.message.FindNoCase( "Invalid row number (65536)" ) )
					Throw( type=exceptionType, message="Too many rows", detail="Binary spreadsheets are limited to 65535 rows. Consider using an XML format spreadsheet instead." );
				else
					rethrow;
			}
		}
		return row;
	}

	private any function createWorkBook(
		required string sheetName
		,boolean xmlFormat=false
		,boolean streamingXml=false
		,numeric streamingWindowSize=100
	){
		validateSheetName( arguments.sheetName );
		if( !arguments.xmlFormat )
			return loadClass( variables.HSSFWorkbookClassName ).init();
		if( arguments.streamingXml ){
			if( ( !IsValid( "integer", arguments.streamingWindowSize ) || arguments.streamingWindowSize < 1 ) )
			Throw( type=exceptionType, message="Invalid 'streamingWindowSize' argument", detail="'streamingWindowSize' must be an integer value greater than 1" );
			return loadClass( variables.SXSSFWorkbookClassName ).init( JavaCast( "int", streamingWindowSize ) );
		}
		return loadClass( variables.XSSFWorkbookClassName ).init();
	}

	private query function deleteHiddenColumnsFromQuery( required sheet, required query result ){
		var startIndex = ( arguments.sheet.totalColumnCount -1 );
		for( var colIndex = startIndex; colIndex GTE 0; colIndex-- ){
			if( !arguments.sheet.object.isColumnHidden( JavaCast( "int", colIndex ) ) )
				continue;
			var columnNumber = ( colIndex +1 );
			arguments.result = _queryDeleteColumn( arguments.result, arguments.sheet.columnNames[ columnNumber ] );
			arguments.sheet.totalColumnCount--;
			arguments.sheet.columnNames.DeleteAt( columnNumber );
		}
		return arguments.result;
	}

	private void function deleteSheetAtIndex( required workbook, required numeric sheetIndex ){
		arguments.workbook.removeSheetAt( JavaCast( "int", arguments.sheetIndex ) );
	}

	private string function detectValueDataType( required value ){
		// Numeric must precede date test
		// Golden default rule: treat numbers with leading zeros as STRINGS: not numbers (lucee) or dates (ACF);
		// Do not detect booleans: leave as strings
		if( REFind( "^0[\d]+", arguments.value ) )
			return "string";
		if( IsNumeric( arguments.value ) )
			return "numeric";
		if( _isDate( arguments.value ) )
			return "date";
		if( !Len( Trim( arguments.value ) ) )
			return "blank";
		return "string";
	}

	private void function downloadBinaryVariable( required binaryVariable, required string filename, required contentType ){
		cfheader( name="Content-Disposition", value='attachment; filename="#arguments.filename#"' );
		cfcontent( type=arguments.contentType, variable="#arguments.binaryVariable#", reset="true" );
	}

	private void function encryptFile( required string filepath, required string password, required string algorithm ){
		/* See http://poi.apache.org/encryption.html */
		/* NB: Not all spreadsheet programs support this type of encryption */
		// set up the encryptor with the chosen algo
		lock name="#arguments.filepath#" timeout=5 {
			var mode = loadClass( "org.apache.poi.poifs.crypt.EncryptionMode" );
			switch( arguments.algorithm ){
				case "agile":
					var info = loadClass( "org.apache.poi.poifs.crypt.EncryptionInfo" ).init( mode.agile );
					break;
				case "standard":
					var info = loadClass( "org.apache.poi.poifs.crypt.EncryptionInfo" ).init( mode.standard );
					break;
				case "binaryRC4":
					var info = loadClass( "org.apache.poi.poifs.crypt.EncryptionInfo" ).init( mode.binaryRC4 );
					break;
			}
			var encryptor = info.getEncryptor();
			encryptor.confirmPassword( JavaCast( "string", arguments.password ) );
			try{
				// set up a POI filesystem object
				var poifs = loadClass( "org.apache.poi.poifs.filesystem.POIFSFileSystem" );
				try{
					// set up an encrypted stream withini the POI filesystem
					var encryptedStream = encryptor.getDataStream( poifs );
					// read in the unencrypted wb file and write it to the encrypted stream
					var workbook = workbookFromFile( arguments.filepath );
					workbook.write( encryptedStream );
				}
				finally{
					// make sure encrypted stream in closed
					if( local.KeyExists( "encryptedStream" ) )
						encryptedStream.close();
				}
				try{
					// write the encrypted POI filesystem to file, replacing the unencypted version
					var outputStream = CreateObject( "java", "java.io.FileOutputStream" ).init( arguments.filepath );
					poifs.writeFilesystem( outputStream );
					outputStream.flush();
				}
				finally{
					// always close the stream. otherwise file may be left in a locked state if an unexpected error occurs
					if( local.KeyExists( "outputStream" ) )
						outputStream.close();
				}
			}
			finally{
				if( local.KeyExists( "poifs" ) )
					poifs.close();
			}
		}
	}

	private numeric function estimateColumnWidth( required workbook, required any value ){
		/* Estimates approximate column width based on cell value and default character width. */
		/*
		"Excel bases its measurement of column widths on the number of digits (specifically, the number of zeros) in the column, using the Normal style font."
		This function approximates the column width using the number of characters and the default character width in the normal font. POI expresses the width in 1/256 of Excel's character unit. The maximum size in POI is: (255 * 256)
		*/
		var defaultWidth = getDefaultCharWidth( arguments.workbook );
		var numOfChars = Len( arguments.value );
		var width = ( numOfChars * defaultWidth +5 ) / ( defaultWidth * 256 );
	    // Do not allow the size to exceed POI's maximum
		return min( width, ( 255 * 256 ) );
	}

	private array function extractRanges( required string rangeList ){
		/*
		A range is a comma-delimited list of ranges, and each value can be either a single number or a range of numbers with a hyphen. Ignores any white space.
		Parses and validates a list of row/column numbers. Returns an array of structures with the keys: startAt, endAt
		*/
		var result = [];
		var rangeTest = "^[0-9]{1,}(-[0-9]{1,})?$";
		var ranges = ListToArray( arguments.rangeList );
		for( var thisRange in ranges ){
			/* remove all white space */
			thisRange.reReplace( "\s+","","ALL" );
			if( !REFind( rangeTest, thisRange ) )
				Throw( type=exceptionType, message="Invalid range value", detail="The range value '#thisRange#' is not valid." );
			var parts = ListToArray( thisRange,"-" );
			//if this is a single number, the start/endAt values are the same
			var range = {
				startAt: parts[ 1 ]
				,endAt: parts[ parts.Len() ]
			};
			result.Append( range );
		}
		return result;
	}

	private string function filenameSafe( required string input ){
		var charsToRemove	=	"\|\\\*\/\:""<>~&";
		var result = arguments.input.reReplace( "[#charsToRemove#]+", "", "ALL" ).Left( 255 );
		if( result.IsEmpty() )
			return "renamed"; // in case all chars have been replaced (unlikely but possible)
		return result;
	}

	private void function doFillMergedCellsWithVisibleValue( required workbook, required sheet ){
		if( !sheetHasMergedRegions( arguments.sheet ) )
			return;
		for( var regionIndex = 0; regionIndex LT arguments.sheet.getNumMergedRegions(); regionIndex++ ){
			var region = arguments.sheet.getMergedRegion( regionIndex );
			var regionStartRowNumber = ( region.getFirstRow() +1 );
			var regionEndRowNumber = ( region.getLastRow() +1 );
			var regionStartColumnNumber = ( region.getFirstColumn() +1 );
			var regionEndColumnNumber = ( region.getLastColumn() +1 );
			var visibleValue = getCellValue( arguments.workbook, regionStartRowNumber, regionStartColumnNumber );
			setCellRangeValue( arguments.workbook, visibleValue, regionStartRowNumber, regionEndRowNumber, regionStartColumnNumber, regionEndColumnNumber );
		}
	}

	private string function generateUniqueSheetName( required workbook ){
		/* Generates a unique sheet name (Sheet1, Sheet2, etecetera). */
		var startNumber = ( arguments.workbook.getNumberOfSheets() +1 );
		var maxRetry = ( startNumber +250 );
		for( var sheetNumber = startNumber; sheetNumber LTE maxRetry; sheetNumber++ ){
			var proposedName = "Sheet" & sheetNumber;
			if( !sheetExists( arguments.workbook, proposedName ) )
				return proposedName;
		}
		/* this should never happen. but if for some reason it did, warn the action failed and abort */
		Throw( type=exceptionType, message="Unable to generate name", detail="Unable to generate a unique sheet name" );
	}

	private any function getActiveSheet( required workbook ){
		return arguments.workbook.getSheetAt( JavaCast( "int", arguments.workbook.getActiveSheetIndex() ) );
	}

	private any function getActiveSheetName( required workbook ){
		return getActiveSheet( arguments.workbook ).getSheetName();
	}

	private numeric function getAWTFontStyle( required any poiFont ){
		var font = loadClass( "java.awt.Font" );
		var isBold = arguments.poiFont.getBold();
		if( isBold && arguments.poiFont.getItalic() )
	  	return BitOr( font.BOLD, font.ITALIC );
		if( isBold )
			return font.BOLD;
		if( arguments.poiFont.getItalic() )
			return font.ITALIC;
		return font.PLAIN;
	}

	private any function getCellAt( required workbook, required numeric rowNumber, required numeric columnNumber ){
		if( !cellExists( argumentCollection=arguments ) )
			Throw( type=exceptionType, message="Invalid cell", detail="The requested cell [#arguments.rowNumber#,#arguments.columnNumber#] does not exist in the active sheet" );
		var rowIndex = ( arguments.rowNumber -1 );
		var columnIndex = ( arguments.columnNumber -1 );
		return getActiveSheet( arguments.workbook ).getRow( JavaCast( "int", rowIndex ) ).getCell( JavaCast( "int", columnIndex ) );
	}

	private any function getCellRangeAddressFromReference( required string rangeReference ){
		/* rangeReference = usually a standard area ref (e.g. "B1:D8"). May be a single cell ref (e.g. "B5") in which case the result is a 1 x 1 cell range. May also be a whole row range (e.g. "3:5"), or a whole column range (e.g. "C:F") */
		return loadClass( "org.apache.poi.ss.util.CellRangeAddress" ).valueOf( JavaCast( "String", arguments.rangeReference ) );
	}

	private any function getCellUtil(){
		if( IsNull( variables.cellUtil ) )
			variables.cellUtil = loadClass( "org.apache.poi.ss.util.CellUtil" );
		return variables.cellUtil;
	}

	private any function getCellValueAsType( required workbook, required cell ){
		/* Get the value of the cell based on the data type. The thing to worry about here is cell forumlas and cell dates. Formulas can be strange and dates are stored as numeric types. Here I will just grab dates as floats and formulas I will try to grab as numeric values. */
		if( cellIsOfType( arguments.cell, "NUMERIC" ) ){
			/* Get numeric cell data. This could be a standard number, could also be a date value. */
			var dateUtil = getDateUtil();
			if( dateUtil.isCellDateFormatted( arguments.cell ) ){
				var cellValue = arguments.cell.getDateCellValue();
				if( DateCompare( "1899-12-31", cellValue, "d" ) EQ 0 ) // TIME
					return getFormatter().formatCellValue( arguments.cell );//return as a time formatted string to avoid default epoch date 1899-12-31
				return cellValue;
			}
			return arguments.cell.getNumericCellValue();
		}
		if( cellIsOfType( arguments.cell, "FORMULA" ) ){
			var formulaEvaluator = arguments.workbook.getCreationHelper().createFormulaEvaluator();
			try{
				return getFormatter().formatCellValue( arguments.cell, formulaEvaluator );
			}
			catch( any exception ){
				Throw( type=exceptionType, message="Failed to run formula", detail="There is a problem with the formula in sheet #arguments.cell.getSheet().getSheetName()# row #( arguments.cell.getRowIndex() +1 )# column #( arguments.cell.getColumnIndex() +1 )#");
			}
		}
		if( cellIsOfType( arguments.cell, "BOOLEAN" ) )
			return arguments.cell.getBooleanCellValue();
	 	if( cellIsOfType( arguments.cell, "BLANK" ) )
			return "";
		try{
			return arguments.cell.getStringCellValue();
		}
		catch( any exception ){
			return "";
		}
	}

	private any function getDateUtil(){
		if( IsNull( variables.dateUtil ) )
			variables.dateUtil = loadClass( "org.apache.poi.ss.usermodel.DateUtil" );
		return variables.dateUtil;
	}

	private string function getDateTimeValueFormat( required any value ){
		/* Returns the default date mask for the given value: DATE (only), TIME (only) or TIMESTAMP */
		var dateTime = ParseDateTime( arguments.value );
		var dateOnly = CreateDate( Year( dateTime ), Month( dateTime ), Day( dateTime ) );
		if( DateCompare( arguments.value, dateOnly, "s" ) EQ 0 )
			return variables.dateFormats.DATE;
		if( DateCompare( "1899-12-30", dateOnly, "d" ) EQ 0 )
			return variables.dateFormats.TIME;
		return variables.dateFormats.TIMESTAMP;
	}

	private numeric function getDefaultCharWidth( required workbook ){
		/* Estimates the default character width using Excel's 'Normal' font */
		/* this is a compromise between hard coding a default value and the more complex method of using an AttributedString and TextLayout */
		var defaultFont = arguments.workbook.getFontAt( 0 );
		var style = getAWTFontStyle( defaultFont );
		var font = loadClass( "java.awt.Font" );
		var javaFont = font.init( defaultFont.getFontName(), style, defaultFont.getFontHeightInPoints() );
		// this works
		var transform = CreateObject( "java", "java.awt.geom.AffineTransform" );
		var fontContext = CreateObject( "java", "java.awt.font.FontRenderContext" ).init( transform, true, true );
		var bounds = javaFont.getStringBounds( "0", fontContext );
		return bounds.getWidth();
	}

	private numeric function getFirstRowNum( required workbook ){
		var sheet = getActiveSheet( arguments.workbook );
		var firstRow = sheet.getFirstRowNum();
		if( firstRow EQ 0 AND sheet.getPhysicalNumberOfRows() EQ 0 )
			return -1;
		return firstRow;
	}

	private any function getFormatter(){
		/* Returns cell formatting utility object ie org.apache.poi.ss.usermodel.DataFormatter */
		if( IsNull( variables.dataFormatter ) )
			variables.dataFormatter = loadClass( "org.apache.poi.ss.usermodel.DataFormatter" ).init();
		return dataFormatter;
	}

	private array function getJarPaths(){
		var libPath = GetDirectoryFromPath( GetCurrentTemplatePath() ) & "lib/";
		return DirectoryList( libPath );
	}

	private struct function getJavaColorRGB( required string colorName ){
		/* Returns a struct containing RGB values from java.awt.Color for the color name passed in */
		var findColor = arguments.colorName.Trim().UCase();
		var color = CreateObject( "Java", "java.awt.Color" );
		if( IsNull( color[ findColor ] ) OR !IsInstanceOf( color[ findColor ], "java.awt.Color" ) )//don't use member functions on color
			Throw( type=exceptionType, message="Invalid color", detail="The color provided (#arguments.colorName#) is not valid." );
		color = color[ findColor ];
		var colorRGB = {
			red: color.getRed()
			,green: color.getGreen()
			,blue: color.getBlue()
		};
		return colorRGB;
	}

	private numeric function getLastRowNum( required workbook, sheet=getActiveSheet( workbook ) ){
		var lastRow = arguments.sheet.getLastRowNum();
		if( lastRow EQ 0 AND arguments.sheet.getPhysicalNumberOfRows() EQ 0 )
			return -1; //The sheet is empty. Return -1 instead of 0
		return lastRow;
	}

	private numeric function getNextEmptyRow( workbook ){
		return ( getLastRowNum( arguments.workbook ) +1 );
	}

	private array function getQueryColumnFormats( required query query ){
		/* extract the query columns and data types  */
		var metadata = GetMetaData( arguments.query );
		/* assign default formats based on the data type of each column */
		for( var col in metadata ){
			var columnType = col.typeName?: "";// typename is missing in ACF if not specified in the query
			switch( columnType ){
				case "DATE": case "TIMESTAMP":
					col.cellDataType = "DATE";
				break;
				case "TIME":
					col.cellDataType = "TIME";
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

	private string function getRgbTripletForStyleColorFormat( required workbook, required cellStyle, required string format ){
		var rgbTriplet = [];
		var isXlsx = isXmlFormat( arguments.workbook );
		var colorObject = "";
		if( !isXlsx )
			var palette = arguments.workbook.getCustomPalette();
		switch( arguments.format ){
			case "bottombordercolor":
				colorObject = isXlsx? arguments.cellStyle.getBottomBorderXSSFColor(): palette.getColor( arguments.cellStyle.getBottomBorderColor() );
				break;
			case "fgcolor":
				colorObject = isXlsx? arguments.cellStyle.getFillForegroundXSSFColor(): palette.getColor( arguments.cellStyle.getFillForegroundColor() );
				break;
			case "leftbordercolor":
				colorObject = isXlsx? arguments.cellStyle.getLeftBorderXSSFColor(): palette.getColor( arguments.cellStyle.getLeftBorderColor() );
				break;
			case "rightbordercolor":
				colorObject = isXlsx? arguments.cellStyle.getRightBorderXSSFColor(): palette.getColor( arguments.cellStyle.getRightBorderColor() );
				break;
			case "topbordercolor":
				colorObject = isXlsx? arguments.cellStyle.getTopBorderXSSFColor(): palette.getColor( arguments.cellStyle.getTopBorderColor() );
				break;
		}
		if( IsNull( colorObject ) OR IsSimpleValue( colorObject) ) // HSSF will return an empty string rather than a null if the color doesn't exist
			return "";
		rgbTriplet = isXlsx? convertSignedRGBToPositiveTriplet( colorObject.getRGB() ): colorObject.getTriplet();
		return ArrayToList( rgbTriplet );
	}

	private array function getRowData( required workbook, required row, array columnRanges=[], boolean includeRichTextFormatting=false ){
		var result = [];
		if( !columnRanges.Len() ){
			var columnRange = {
				startAt: 1
				,endAt: arguments.row.GetLastCellNum()
			};
			arguments.columnRanges = [ columnRange ];
		}
		for( var thisRange in arguments.columnRanges ){
			for( var i = thisRange.startAt; i LTE thisRange.endAt; i++ ){
				var colIndex = ( i-1 );
				var cell = arguments.row.GetCell( JavaCast( "int", colIndex ) );
				if( IsNull( cell ) ){
					result.Append( "" );
					continue;
				}
				var cellValue = getCellValueAsType( arguments.workbook, cell );
				if( arguments.includeRichTextFormatting AND cellIsOfType( cell, "STRING" ) )
					cellValue = richStringCellValueToHtml( arguments.workbook, cell,cellValue );
				result.Append( cellValue );
			}
		}
		return result;
	}

	private any function getSheetByName( required workbook, required string sheetName ){
		validateSheetExistsWithName( arguments.workbook, arguments.sheetName );
		return arguments.workbook.getSheet( JavaCast( "string", arguments.sheetName ) );
	}

	private any function getSheetByNameOrNumber( required workbook, string sheetName, numeric sheetNumber ){
		var sheetNameSupplied = ( arguments.KeyExists( "sheetName" ) AND Len( arguments.sheetName ) );
		if( sheetNameSupplied AND arguments.KeyExists( "sheetNumber" ) )
			Throw( type=exceptionType, message="Invalid arguments", detail="Specify either a sheetName or sheetNumber, not both" );
		if( sheetNameSupplied )
			return getSheetByName( arguments.workbook, arguments.sheetName );
		if( arguments.KeyExists( "sheetNumber" ) )
			return getSheetByNumber( arguments.workbook, arguments.sheetNumber );
		return getActiveSheet( arguments.workbook );
	}

	private any function getSheetByNumber( required workbook, required numeric sheetNumber ){
		validateSheetNumber( arguments.workbook, arguments.sheetNumber );
		var sheetIndex = ( arguments.sheetNumber -1 );
		return arguments.workbook.getSheetAt( sheetIndex );
	}

	private numeric function getSheetIndexFromName( required workbook, required string sheetName ){
		//returns -1 if non-existent
		return arguments.workbook.getSheetIndex( JavaCast( "string", arguments.sheetName ) );
	}

	private string function getUnderlineFormatAsString( required cellFont ){
		var lookup = {};
		lookup[ 0 ] = "none";
		lookup[ 1 ] = "single";
		lookup[ 2 ] = "double";
		lookup[ 33 ] = "single accounting";
		lookup[ 34 ] = "double accounting";
		if( lookup.KeyExists( arguments.cellFont.getUnderline() ) )
			return lookup[ arguments.cellFont.getUnderline() ];
		return "unknown";
	}

	private void function handleInvalidSpreadsheetFile( required string path ){
		var detail = "The file #arguments.path# does not appear to be a binary or xml spreadsheet.";
		if( isCsvOrTextFile( arguments.path ) )
			detail &= " It may be a CSV file, in which case use 'csvToQuery()' to read it";
		Throw( type="cfsimplicity.lucee.spreadsheet.invalidFile", message="Invalid spreadsheet file", detail=detail );
	}

	private any function initializeCell( required workbook, required numeric rowNumber, required numeric columnNumber ){
		var rowIndex = JavaCast( "int", ( arguments.rowNumber -1 ) );
		var columnIndex = JavaCast( "int", ( arguments.columnNumber -1 ) );
		var rowObject = getCellUtil().getRow( rowIndex, getActiveSheet( arguments.workbook ) );
		var cellObject = getCellUtil().getCell( rowObject, columnIndex );
		return cellObject;
	}

	private boolean function isCsvOrTextFile( required string path ){
		var contentType = FileGetMimeType( arguments.path ).ListLast( "/" );
		return ListFindNoCase( "plain,csv", contentType );//Lucee=text/plain ACF=text/csv
	}

	private boolean function isDateObject( required input ){
		return arguments.input.getClass().getName() IS "java.util.Date";
	}

	private boolean function isString( required input ){
		return arguments.input.getClass().getName() IS "java.lang.String";
	}

	private function loadClass( required string javaclass ){
		if( !requiresJavaLoader ){
			// If not using JL, *the correct* POI jars must be in the class path and any older versions *removed*
			try{
				javaClassesLastLoadedVia = "The java class path";
				return CreateObject( "java", arguments.javaclass );
			}
			catch( any exception ){
				javaClassesLastLoadedVia = "JavaLoader";
				return loadClassUsingJavaLoader( arguments.javaclass );
			}
		}
		javaClassesLastLoadedVia = "JavaLoader";
		return loadClassUsingJavaLoader( arguments.javaclass );
	}

	private function loadClassUsingJavaLoader( required string javaclass ){
		if( !server.KeyExists( javaLoaderName ) )
			server[ javaLoaderName ] = CreateObject( "component", javaLoaderDotPath ).init( loadPaths=getJarPaths(), loadColdFusionClassPath=false, trustedSource=true );
		return server[ javaLoaderName ].create( arguments.javaclass );
	}

	private void function moveSheet( required workbook, required string sheetName, required string moveToIndex ){
		arguments.workbook.setSheetOrder( JavaCast( "String", arguments.sheetName ), JavaCast( "int", arguments.moveToIndex ) );
	}

	private array function parseRowData( required string line, required string delimiter, boolean handleEmbeddedCommas=true ){
		var elements = ListToArray( arguments.line, arguments.delimiter );
		var potentialQuotes = 0;
		arguments.line = ToString( arguments.line );
		if( arguments.delimiter EQ "," AND arguments.handleEmbeddedCommas )
			potentialQuotes = arguments.line.ReplaceAll( "[^']", "" ).length();
		if( potentialQuotes <= 1 )
		  return elements;
		//For ACF compatibility, find any values enclosed in single quotes and treat them as a single element.
		var currentValue = 0;
		var nextValue = "";
		var isEmbeddedValue = false;
		var values = [];
		var buffer = CreateObject( "Java", "java.lang.StringBuilder" ).init();
		var maxElements = ArrayLen( elements );

		for( var i = 1; i LTE maxElements; i++ ) {
		  currentValue = Trim( elements[ i ] );
		  nextValue = i < maxElements ? elements[ i + 1 ] : "";
		  var isComplete = false;
		  var hasLeadingQuote = ( currentValue.Left( 1 ) IS "'" );
		  var hasTrailingQuote = ( currentValue.Right( 1 ) IS "'" );
		  var isFinalElement = ( i == maxElements );
		  if( hasLeadingQuote )
			  isEmbeddedValue = true;
		  if( isEmbeddedValue AND hasTrailingQuote )
			  isComplete = true;
		  /* We are finished with this value if:
			  * no quotes were found OR
			  * it is the final value OR
			  * the next value is embedded in quotes
		  */
		  if( !isEmbeddedValue || isFinalElement || ( nextValue.Left( 1 ) IS "'" ) )
			  isComplete = true;
		  if( isEmbeddedValue || isComplete ){
			  // if this a partial value, append the delimiter
			  if( isEmbeddedValue AND buffer.length() GT 0 )
				  buffer.Append( "," );
			  buffer.Append( elements[ i ] );
		  }
		  if( isComplete ){
			  var finalValue = buffer.toString();
			  var startAt = finalValue.indexOf( "'" );
			  var endAt = finalValue.lastIndexOf( "'" );
			  if( isEmbeddedValue AND startAt GTE 0 AND endAt GT startAt )
				  finalValue = finalValue.substring( ( startAt +1 ), endAt );
			  values.Append( finalValue );
			  buffer.setLength( 0 );
			  isEmbeddedValue = false;
		  }
	  }
	  return values;
	}

	private string function queryToCsv( required query query, numeric headerRow, boolean includeHeaderRow=false ){
		var result = CreateObject( "Java", "java.lang.StringBuilder" ).init();
		var crlf = Chr( 13 ) & Chr( 10 );
		var columns = _queryColumnArray( arguments.query );
		var generateHeaderRow = ( arguments.includeHeaderRow && arguments.KeyExists( "headerRow" ) && Val( arguments.headerRow ) );
		if( generateHeaderRow )
			result.Append( generateCsvRow( columns ) );
		for( var row in arguments.query ){
			var rowValues = [];
			for( var column in columns )
				rowValues.Append( row[ column ] );
			result.Append( crlf & generateCsvRow( rowValues ) );
		}
		return result.toString().Trim();
	}

	private string function generateCsvRow( required array values, delimiter="," ){
		var result = CreateObject( "Java", "java.lang.StringBuilder" ).init();
		for( var value in arguments.values ){
			if( isDateObject( value ) )
				value = DateTimeFormat( value, dateFormats.DATETIME );
			value = Replace( value, '"', '""', "ALL" );//can't use member function in case its a non-string
			result.Append( '#arguments.delimiter#"#value#"' );
		}
		return result.toString().substring( 1 );
	}

	private string function queryToHtml( required query query, numeric headerRow, boolean includeHeaderRow=false ){
		var result = CreateObject( "Java", "java.lang.StringBuilder" ).init();
		var columns = _queryColumnArray( arguments.query );
		var generateHeaderRow = ( arguments.includeHeaderRow && arguments.KeyExists( "headerRow" ) && Val( arguments.headerRow ) );
		if( generateHeaderRow ){
			result.Append( "<thead>" );
			result.Append( generateHtmlRow( columns, true ) );
			result.Append( "</thead>" );
		}
		result.Append( "<tbody>" );
		for( var row in arguments.query ){
			var rowValues = [];
			for( var column in columns )
				rowValues.Append( row[ column ] );
			result.Append( generateHtmlRow( rowValues ) );
		}
		result.Append( "</tbody>" );
		return result.toString();
	}

	private string function generateHtmlRow( required array values, boolean isHeader=false ){
		var result = CreateObject( "Java", "java.lang.StringBuilder" ).init();
		result.Append( "<tr>" );
		var columnTag = arguments.isHeader? "th": "td";
		for( var value in arguments.values ){
			if( isDateObject( value ) )
				value = DateTimeFormat( value, dateFormats.DATETIME );
			result.Append( "<#columnTag#>#value#</#columnTag#>" );
		}
		result.Append( "</tr>" );
		return result.toString();
	}

	private boolean function rowIsEmpty( required row ){
		for( var i = arguments.row.getFirstCellNum(); i LT arguments.row.getLastCellNum(); i++ ){
	    var cell = arguments.row.getCell( i );
	    if( !IsNull( cell ) && !cellIsOfType( cell, "BLANK" ) )
	      return false;
	  }
	  return true;
	}

	private void function setCellValueAsType( required workbook, required cell, required value, string type ){
		if( !arguments.KeyExists( "type" ) ) //autodetect type
			arguments.type = detectValueDataType( arguments.value );
		else if( !ListFindNoCase( "string,numeric,date,boolean,blank", arguments.type ) )
			Throw( type=exceptionType, message="Invalid data type: '#arguments.type#'", detail="The data type must be one of 'string', 'numeric', 'date' 'boolean' or 'blank'." );
		/* Note: To properly apply date/number formatting:
			- cell type must be CELL_TYPE_NUMERIC
			- cell value must be applied as a java.util.Date or java.lang.Double (NOT as a string)
			- cell style must have a dataFormat (datetime values only)
 		*/
		switch( arguments.type ){
			case "numeric":
				arguments.cell.setCellType( arguments.cell.CellType.NUMERIC );
				arguments.cell.setCellValue( JavaCast( "double", Val( arguments.value ) ) );
				return;
			case "date":
				//handle empty strings which can't be treated as dates
				if( !Len( Trim( arguments.value ) ) ){
					arguments.cell.setCellType( arguments.cell.CellType.BLANK ); //no need to set the value: it will be blank
					return;
				}
				var cellFormat = getDateTimeValueFormat( arguments.value );
				var formatter = arguments.workbook.getCreationHelper().createDataFormat();
				//Use setCellStyleProperty() which will try to re-use an existing style rather than create a new one for every cell which may breach the 4009 styles per wookbook limit
				getCellUtil().setCellStyleProperty( arguments.cell, getCellUtil().DATA_FORMAT, formatter.getFormat( JavaCast( "string", cellFormat ) ) );
				cell.setCellType( arguments.cell.CellType.NUMERIC );
				/*  Excel's uses a different epoch than CF (1900-01-01 versus 1899-12-30). "Time" only values will not display properly without special handling - */
				if( cellFormat EQ variables.dateFormats.TIME ){
					var dateUtil = getDateUtil();
					arguments.value = TimeFormat( arguments.value, "HH:MM:SS" );
				 	arguments.cell.setCellValue( dateUtil.convertTime( arguments.value ) );
				}
				else
					arguments.cell.setCellValue( ParseDateTime( arguments.value ) );
				return;
			case "boolean":
				//handle empty strings/nulls which can't be treated as booleans
				if( !Len( Trim( arguments.value ) ) ){
					arguments.cell.setCellType( arguments.cell.CellType.BLANK ); //no need to set the value: it will be blank
					return;
				}
				arguments.cell.setCellType( arguments.cell.CellType.BOOLEAN );
				arguments.cell.setCellValue( JavaCast( "boolean", arguments.value ) );
				return;
			case "blank":
				arguments.cell.setCellType( arguments.cell.CellType.BLANK ); //no need to set the value: it will be blank
				return;
		}
		// string cellStyle.getAlignmentEnum()[ JavaCast( "string", UCase( settingValue ) ) ];
		arguments.cell.setCellType( arguments.cell.CellType.STRING );
		arguments.cell.setCellValue( JavaCast( "string", arguments.value ) );
	}

	private boolean function sheetExists( required workbook, string sheetName, numeric sheetNumber ){
		validateSheetNameOrNumberWasProvided( argumentCollection=arguments );
		if( arguments.KeyExists( "sheetName" ) )
			arguments.sheetNumber = ( getSheetIndexFromName( arguments.workbook, arguments.sheetName ) +1 );
			//the position is valid if it an integer between 1 and the total number of sheets in the workbook
		if( arguments.sheetNumber AND ( arguments.sheetNumber EQ Round( arguments.sheetNumber ) ) AND ( arguments.sheetNumber LTE arguments.workbook.getNumberOfSheets() ) )
			return true;
		return false;
	}

	private boolean function sheetHasMergedRegions( required sheet ){
		return ( arguments.sheet.getNumMergedRegions() GT 0 );
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
		var sheet = {
			includeHeaderRow: arguments.includeHeaderRow
			,hasHeaderRow: ( arguments.KeyExists( "headerRow" ) AND Val( arguments.headerRow ) )
			,includeBlankRows: arguments.includeBlankRows
			,columnNames: []
			,columnRanges: []
			,totalColumnCount: 0
		};
		sheet.headerRowIndex = sheet.hasHeaderRow? ( arguments.headerRow -1 ): -1;
		if( arguments.KeyExists( "columns" ) ){
			sheet.columnRanges = extractRanges( arguments.columns );
			sheet.totalColumnCount = columnCountFromRanges( sheet.columnRanges );
		}
		if( arguments.KeyExists( "sheetName" ) ){
			validateSheetExistsWithName( arguments.workbook, arguments.sheetName );
			arguments.sheetNumber = ( getSheetIndexFromName( arguments.workbook, arguments.sheetName ) +1 );
		}
		sheet.object = getSheetByNumber( arguments.workbook, arguments.sheetNumber );
		if( arguments.fillMergedCellsWithVisibleValue )
			doFillMergedCellsWithVisibleValue( arguments.workbook,sheet.object );
		sheet.data = [];
		if( arguments.KeyExists( "rows" ) ){
			var allRanges = extractRanges( arguments.rows );
			for( var thisRange in allRanges ){
				for( var rowNumber = thisRange.startAt; rowNumber LTE thisRange.endAt; rowNumber++ ){
					var rowIndex = ( rowNumber -1 );
					addRowToSheetData( arguments.workbook, sheet, rowIndex, arguments.includeRichTextFormatting );
				}
			}
		}
		else {
			var lastRowIndex = sheet.object.GetLastRowNum();// zero based
			for( var rowIndex = 0; rowIndex LTE lastRowIndex; rowIndex++ )
				addRowToSheetData( arguments.workbook, sheet, rowIndex, arguments.includeRichTextFormatting );
		}
		//generate the query columns
		if( arguments.KeyExists( "columnNames" ) AND arguments.columnNames.Len() ){
			arguments.columnNames = arguments.columnNames.ListToArray();
			var specifiedColumnCount = arguments.columnNames.Len();
			for( var i = 1; i LTE sheet.totalColumnCount; i++ ){
				// IsNull/IsDefined doesn't work.
				var columnName = arguments.columnNames[ i ]?: "column" & i;
				sheet.columnNames.Append( columnName );
			}
		}
		else if( sheet.hasHeaderRow ){
			var headerRowObject = sheet.object.getRow( JavaCast( "int", sheet.headerRowIndex ) );
			var rowData = getRowData( arguments.workbook, headerRowObject, sheet.columnRanges );
			var i = 1;
			for( var value in rowData ){
				var columnName = "column" & i;
				if( isString( value ) AND value.Len() )
					columnName = value;
				sheet.columnNames.Append( columnName );
				i++;
			}
		}
		else {
			for( var i=1; i LTE sheet.totalColumnCount; i++ )
				sheet.columnNames.Append( "column" & i );
		}
		var result = _queryNew( sheet.columnNames, "", sheet.data );
		if( !arguments.includeHiddenColumns ){
			result = deleteHiddenColumnsFromQuery( sheet, result );
			if( sheet.totalColumnCount EQ 0 )
				return QueryNew( "" );// all columns were hidden: return a blank query.
		}
		return result;
	}

	private void function toggleColumnHidden( required workbook, required numeric columnNumber, required boolean state ){
		getActiveSheet( arguments.workbook ).setColumnHidden( JavaCast( "int", arguments.columnNumber-1 ), JavaCast( "boolean", arguments.state ) );
	}

	private void function toggleRowHidden( required workbook, required numeric row, required boolean state ){
		var rowIndex = ( arguments.row -1 );
		getActiveSheet( arguments.workbook ).getRow( JavaCast( "int", rowIndex ) ).setZeroHeight( JavaCast( "boolean", arguments.state ) );
	}

	private void function validateSheetExistsWithName( required workbook, required string sheetName ){
		if( !sheetExists( workbook=arguments.workbook, sheetName=arguments.sheetName ) )
			Throw( type=exceptionType, message="Invalid sheet name [#arguments.sheetName#]", detail="The specified sheet was not found in the current workbook." );
	}

	private void function validateSheetNumber( required workbook, required numeric sheetNumber ){
		if( !sheetExists( workbook=arguments.workbook, sheetNumber=arguments.sheetNumber ) ){
			var sheetCount = arguments.workbook.getNumberOfSheets();
			Throw( type=exceptionType, message="Invalid sheet number [#arguments.sheetNumber#]", detail="The sheetNumber must a whole number between 1 and the total number of sheets in the workbook [#sheetCount#]" );
		}
	}

	private void function validateSheetName( required string sheetName ){
		var characterCount = Len( arguments.sheetName );
		if( characterCount GT 31 )
			Throw( type=exceptionType, message="Invalid sheet name", detail="The sheetname contains too many characters [#characterCount#]. The maximum is 31." );
		var poiTool = loadClass( "org.apache.poi.ss.util.WorkbookUtil" );
		try{
			poiTool.validateSheetName( JavaCast( "String", arguments.sheetName ) );
		}
		catch( "java.lang.IllegalArgumentException" exception ){
			Throw( type=exceptionType, message="Invalid characters in sheet name", detail=exception.message );
		}
		catch( "java.lang.reflect.InvocationTargetException" exception ){
			//ACF
			Throw( type=exceptionType, message="Invalid characters in sheet name", detail=exception.message );
		}
	}

	private void function validateSheetNameOrNumberWasProvided(){
		if( !arguments.KeyExists( "sheetName" ) AND !arguments.KeyExists( "sheetNumber" ) )
			Throw( type=exceptionType, message="Missing Required Argument", detail="Either sheetName or sheetNumber must be provided" );
		if( arguments.KeyExists( "sheetName" ) AND arguments.KeyExists( "sheetNumber" ) )
			Throw( type=exceptionType, message="Too Many Arguments", detail="Only one argument is allowed. Specify either a sheetName or sheetNumber, not both" );
	}

	private any function workbookFromFile( required string path, string password ){
		// works with both xls and xlsx
		try{
			lock name="#arguments.path#" timeout=5 {
				var className = "org.apache.poi.ss.usermodel.WorkbookFactory";
				var file = CreateObject( "java", "java.io.FileInputStream" ).init( arguments.path );
				if( arguments.KeyExists( "password" ) )
					return loadClass( className ).create( file, arguments.password );
				return loadClass( className ).create( file );
			}
		}
		catch( org.apache.poi.openxml4j.exceptions.InvalidFormatException exception ){
			handleInvalidSpreadsheetFile( arguments.path );
		}
		catch( org.apache.poi.hssf.OldExcelFormatException exception ){
			Throw( type="cfsimplicity.lucee.spreadsheet.OldExcelFormatException", message="Invalid spreadsheet format", detail="The file #arguments.path# was saved in a format that is too old. Please save it as an 'Excel 97/2000/XP' file or later." );
		}
		catch( any exception ){
			//For ACF which doesn't return the correct exception types
			if( exception.message CONTAINS "Your InputStream was neither" )
				handleInvalidSpreadsheetFile( arguments.path );
			if( exception.message CONTAINS "spreadsheet seems to be Excel 5" )
				Throw( type="cfsimplicity.lucee.spreadsheet.OldExcelFormatException", message="Invalid spreadsheet format", detail="The file #arguments.path# was saved in a format that is too old. Please save it as an 'Excel 97/2000/XP' file or later." );
			rethrow;
		}
		finally{
			if( local.KeyExists( "file" ) )
				file.close();
		}
	}

	private struct function xmlInfo( required workbook ){
		var documentProperties = arguments.workbook.getProperties().getExtendedProperties().getUnderlyingProperties();
		var coreProperties = arguments.workbook.getProperties().getCoreProperties();
		var result = {
			author: coreProperties.getCreator()?:""
			,category: coreProperties.getCategory()?:""
			,comments: coreProperties.getDescription()?:""
			,creationDate: coreProperties.getCreated()?:""
			,lastEdited: coreProperties.getModified()?:""
			,subject: coreProperties.getSubject()?:""
			,title: coreProperties.getTitle()?:""
			,keywords: coreProperties.getKeywords()?:""
			,lastSaved: ""// not available in xml
			,manager: documentProperties.getManager()?:""
			,company: documentProperties.getCompany()?:""
		};
		// lastAuthor is a java.util.Option object with different behaviour
		if( coreProperties.getUnderlyingProperties().getLastModifiedByProperty().isPresent() )
			result.lastAuthor = coreProperties.getUnderlyingProperties().getLastModifiedByProperty().get();
		return result;
	}

	/* Formatting */

	private string function richStringCellValueToHtml( required workbook, required cell, required cellValue ){
		var richTextValue = arguments.cell.getRichStringCellValue();
		var totalRuns = richTextValue.numFormattingRuns();
		var baseFont = arguments.cell.getCellStyle().getFont( arguments.workbook );
		if( totalRuns EQ 0  )
			return baseFontToHtml( arguments.workbook, arguments.cellValue, baseFont );
		// Runs never start at the beginning: the string before the first run is always in the baseFont format
		var startOfFirstRun = richTextValue.getIndexOfFormattingRun( 0 );
		var initialContents = arguments.cellValue.Mid( 1, startOfFirstRun );//before the first run
		var initialHtml = baseFontToHtml( arguments.workbook, initialContents, baseFont );
		var result = CreateObject( "Java", "java.lang.StringBuilder" ).init();
		result.Append( initialHtml );
		var endOfCellValuePosition = arguments.cellValue.Len();
		for( var runIndex = 0; runIndex LT totalRuns; runIndex++ ){
			var run = {};
			run.index = runIndex;
			run.number = ( runIndex +1 );
			run.font = arguments.workbook.getFontAt( richTextValue.getFontOfFormattingRun( runIndex ) );
			run.css = runFontToHtml( arguments.workbook, baseFont, run.font );
			run.isLast = ( run.number EQ totalRuns );
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

	private string function runFontToHtml( required workbook, required baseFont, required runFont ){
		/* NB: the order of processing is important for the tests to match */
		var cssStyles = CreateObject( "Java", "java.lang.StringBuilder" ).init();
		/* bold */
		if( compare( arguments.runFont.getBold(), arguments.baseFont.getBold() ) )
			cssStyles.Append( fontStyleToCss( "bold", arguments.runFont.getBold() ) );
		/* color */
		if( compare( arguments.runFont.getColor(), arguments.baseFont.getColor() ) AND !fontColorIsBlack( arguments.runFont.getColor() ) )
			cssStyles.Append( fontStyleToCss( "color", arguments.runFont.getColor(), arguments.workbook ) );
		/* italic */
		if( compare( arguments.runFont.getItalic(), arguments.baseFont.getItalic() ) )
			cssStyles.Append( fontStyleToCss( "italic", arguments.runFont.getItalic() ) );
		/* underline/strike */
		if( compare( arguments.runFont.getStrikeout(), arguments.baseFont.getStrikeout() ) OR Compare( arguments.runFont.getUnderline(), arguments.baseFont.getUnderline() ) ){
			var decorationValue	=	[];
			if( !arguments.baseFont.getStrikeout() AND arguments.runFont.getStrikeout() )
				decorationValue.Append( "line-through" );
			if( !arguments.baseFont.getUnderline() AND arguments.runFont.getUnderline() )
				decorationValue.Append( "underline" );
			//if either or both are in the base format, and either or both are NOT in the run format, set the decoration to none.
			if(
					( arguments.baseFont.getUnderline() OR arguments.baseFont.getStrikeout() )
					AND
					( !arguments.runFont.getUnderline() OR !arguments.runFont.getUnderline() )
				){
				cssStyles.Append( fontStyleToCss( "decoration", "none" ) );
			}
			else
				cssStyles.Append( fontStyleToCss( "decoration", decorationValue.ToList( " " ) ) );
		}
		return cssStyles.toString();
	}

	private string function baseFontToHtml( required workbook, required contents, required baseFont ){
		/* the order of processing is important for the tests to match */
		/* font family and size not parsed here because all cells would trigger formatting of these attributes: defaults can't be assumed */
		var cssStyles = CreateObject( "Java", "java.lang.StringBuilder" ).init();
		/* bold */
		if( arguments.baseFont.getBold() )
			cssStyles.Append( fontStyleToCss( "bold", true ) );
		/* color */
		if( !fontColorIsBlack( arguments.baseFont.getColor() ) )
			cssStyles.Append( fontStyleToCss( "color", arguments.baseFont.getColor(), arguments.workbook ) );
		/* italic */
		if( arguments.baseFont.getItalic() )
			cssStyles.Append( fontStyleToCss( "italic", true ) );
		/* underline/strike */
		if( arguments.baseFont.getStrikeout() OR arguments.baseFont.getUnderline() ){
			var decorationValue	=	[];
			if( arguments.baseFont.getStrikeout() )
				decorationValue.Append( "line-through" );
			if( arguments.baseFont.getUnderline() )
				decorationValue.Append( "underline" );
			cssStyles.Append( fontStyleToCss( "decoration", decorationValue.ToList( " " ) ) );
		}
		cssStyles = cssStyles.toString();
		if( cssStyles.IsEmpty() )
			return contents;
		return "<span style=""#cssStyles#"">#arguments.contents#</span>";
	}

	private string function fontStyleToCss( required string styleType, required any styleValue, workbook ){
		/*
		Support limited to:
			bold
			color
			italic
			strikethrough
			single underline
		*/
		switch( arguments.styleType ){
			case "bold":
				return "font-weight:" & ( arguments.styleValue? "bold;": "normal;" );
			case "color":
				if( !arguments.KeyExists( "workbook" ) )
					Throw( type=exceptionType, message="The 'workbook' argument is required when generating color css styles" );
				//http://ragnarock99.blogspot.co.uk/2012/04/getting-hex-color-from-excel-cell.html
				var rgb = arguments.workbook.getCustomPalette().getColor( arguments.styleValue ).getTriplet();
				var javaColor = CreateObject( "Java", "java.awt.Color" ).init( JavaCast( "int", rgb[ 1 ] ), JavaCast( "int", rgb[ 2 ] ), JavaCast( "int", rgb[ 3 ] ) );
				var hex	=	CreateObject( "Java", "java.lang.Integer" ).toHexString( javaColor.getRGB() );
				hex = hex.subString( 2, hex.length() );
				return "color:##" & hex & ";";
			case "italic":
				return "font-style:" & ( arguments.styleValue? "italic;": "normal;" );
			case "decoration":
				return "text-decoration:#arguments.styleValue#;";//need to pass desired combination of "underline" and "line-through"
		}
		Throw( type=exceptionType, message="Unrecognised style for css conversion" );
	}

	private boolean function fontColorIsBlack( required fontColor ){
		return ( arguments.fontColor IS 8 ) OR ( arguments.fontColor IS 32767 );
	}

	private any function buildCellStyle( required workbook, required struct format ){
		/*  TODO: Reuse styles  */
		var cellStyle = arguments.workbook.createCellStyle();
		var formatter = arguments.workbook.getCreationHelper().createDataFormat();
		var font = 0;
		var formatIndex = 0;
		for( var setting in arguments.format ){
			var settingValue = arguments.format[ setting ];
			switch( setting ){
				case "alignment":
					var alignment = cellStyle.getAlignmentEnum()[ JavaCast( "string", UCase( settingValue ) ) ];
					cellStyle.setAlignment( alignment );
				break;
				case "bold":
					font = cloneFont( arguments.workbook, arguments.workbook.getFontAt( cellStyle.getFontIndex() ) );
					font.setBold( JavaCast( "boolean", settingValue ) );
					cellStyle.setFont( font );
				break;
				case "bottomborder":
					var borderStyle = cellStyle.getBorderBottomEnum()[ JavaCast( "string", UCase( settingValue ) ) ];
					cellStyle.setBorderBottom( borderStyle );
				break;
				case "bottombordercolor":
					cellStyle.setBottomBorderColor( getColor( arguments.workbook, settingValue ) );
				break;
				case "color":
					font = cloneFont( arguments.workbook, arguments.workbook.getFontAt( cellStyle.getFontIndex() ) );
					font.setColor( getColor( arguments.workbook, settingValue ) );
					cellStyle.setFont( font );
				break;
				case "dataformat":
					cellStyle.setDataFormat( formatter.getFormat( JavaCast( "string", settingValue ) ) );
				break;
				case "fgcolor":
					cellStyle.setFillForegroundColor( getColor( arguments.workbook, settingValue ) );
					/*  make sure we always apply a fill pattern or the color will not be visible  */
					if( !arguments.KeyExists( "fillpattern" ) ){
						var fillpattern = cellStyle.getFillPatternEnum()[ JavaCast( "string", "SOLID_FOREGROUND" ) ];
						cellStyle.setFillPattern( fillpattern );
					}
				break;
				case "fillpattern":
					if( settingValue IS "nofill" ) //CF 9 docs list "nofill" as opposed to "no_fill"
						settingValue = "NO_FILL";
					var fillpattern = cellStyle.getFillPatternEnum()[ JavaCast( "string", UCase( settingValue ) ) ];
					cellStyle.setFillPattern( fillpattern );
				break;
				case "font":
					font = cloneFont( arguments.workbook, arguments.workbook.getFontAt( cellStyle.getFontIndex() ) );
					font.setFontName( JavaCast( "string", settingValue ) );
					cellStyle.setFont( font );
				break;
				case "fontsize":
					font = cloneFont( arguments.workbook, arguments.workbook.getFontAt( cellStyle.getFontIndex() ) );
					font.setFontHeightInPoints( JavaCast( "int", settingValue ) );
					cellStyle.setFont( font );
				break;
				/*  TODO: Doesn't seem to do anything */
				case "hidden":
					cellStyle.setHidden( JavaCast( "boolean", settingValue ) );
				break;
				case "indent":
					// Only seems to work on MS Excel. XLS limit is 15.
					var indentValue = isXmlFormat( arguments.workbook )? settingValue: Min( 15, settingValue );
					cellStyle.setIndention( JavaCast( "int", indentValue ) );
				break;
				case "italic":
					font = cloneFont( arguments.workbook, arguments.workbook.getFontAt( cellStyle.getFontIndex ( ) ) );
					font.setItalic( JavaCast( "boolean", settingValue ) );
					cellStyle.setFont( font );
				break;
				case "leftborder":
					var borderStyle = cellStyle.getBorderLeftEnum()[ JavaCast( "string", UCase( settingValue ) ) ];
					cellStyle.setBorderLeft( borderStyle );
				break;
				case "leftbordercolor":
					cellStyle.setLeftBorderColor( getColor( arguments.workbook, settingValue ) );
				break;
				/*  TODO: Doesn't seem to do anything */
				case "locked":
					cellStyle.setLocked( JavaCast( "boolean", settingValue ) );
				break;
				case "quoteprefixed":
					cellStyle.setQuotePrefixed( JavaCast( "boolean", settingValue ) );
				break;
				case "rightborder":
					var borderStyle = cellStyle.getBorderRightEnum()[ JavaCast( "string", UCase( settingValue ) ) ];
					cellStyle.setBorderRight( borderStyle );
				break;
				case "rightbordercolor":
					cellStyle.setRightBorderColor( getColor( arguments.workbook, settingValue ) );
				break;
				case "rotation":
					cellStyle.setRotation( JavaCast( "int", settingValue ) );
				break;
				case "strikeout":
					font = cloneFont( arguments.workbook, arguments.workbook.getFontAt( cellStyle.getFontIndex() ) );
					font.setStrikeout( JavaCast( "boolean", settingValue ) );
					cellStyle.setFont( font );
				break;
				case "textwrap":
					cellStyle.setWrapText( JavaCast( "boolean", settingValue ) );
				break;
				case "topborder":
					var borderStyle = cellStyle.getBorderTopEnum()[ JavaCast( "string", UCase( settingValue ) ) ];
					cellStyle.setBorderTop( borderStyle );
				break;
				case "topbordercolor":
					cellStyle.setTopBorderColor( getColor( arguments.workbook, settingValue ) );
				break;
				case "underline":
					var underlineType = 0;
					switch( settingValue ){
						case "none": underlineType = 0;
							break;
						case "single": underlineType = 1;
							break;
						case "double": underlineType = 2;
							break;
						case "single accounting": underlineType = 33;
							break;
						case "double accounting": underlineType = 34;
							break;
						default:
							if( !IsBoolean( settingValue ) )
								return cellStyle; //invalid - do nothing
							underlineType = settingValue? 1: 0;
					}
					font = cloneFont( arguments.workbook, arguments.workbook.getFontAt( cellStyle.getFontIndex() ) );
					font.setUnderline( JavaCast( "byte", underlineType ) );
					cellStyle.setFont( font );
				break;
				case "verticalalignment":
					var alignment = cellStyle.getVerticalAlignmentEnum()[ JavaCast( "string", UCase( settingValue ) ) ];
					cellStyle.setVerticalAlignment( alignment );
				break;
			}
		}
		return cellStyle;
	}

	private any function cloneFont( required workbook, required fontToClone ){
		var newFont = arguments.workbook.createFont();
		/*  copy the existing cell's font settings to the new font  */
		newFont.setBold( arguments.fontToClone.getBold() );
		newFont.setCharSet( arguments.fontToClone.getCharSet() );
		// xlsx fonts contain XSSFColor objects which may have been set as RGB
		newFont.setColor( isXmlFormat( arguments.workbook )? arguments.fontToClone.getXSSFColor(): arguments.fontToClone.getColor() );
		newFont.setFontHeight( arguments.fontToClone.getFontHeight() );
		newFont.setFontName( arguments.fontToClone.getFontName() );
		newFont.setItalic( arguments.fontToClone.getItalic() );
		newFont.setStrikeout( arguments.fontToClone.getStrikeout() );
		newFont.setTypeOffset( arguments.fontToClone.getTypeOffset() );
		newFont.setUnderline( arguments.fontToClone.getUnderline() );
		return newFont;
	}

	private numeric function getColorIndex( required string colorName ){
		var findColor = arguments.colorName.Trim().UCase();
		//check for 9 extra colours from old org.apache.poi.ss.usermodel.IndexedColors and map
		var deprecatedNames = [ "BLACK1", "WHITE1", "RED1", "BRIGHT_GREEN1", "BLUE1", "YELLOW1", "PINK1", "TURQUOISE1", "LIGHT_TURQUOISE1" ];
		if( deprecatedNames.Find( findColor ) )
			findColor = findColor.Left( findColor.Len() - 1 );
		var indexedColors = loadClass( "org.apache.poi.hssf.util.HSSFColor$HSSFColorPredefined" );
		try{
			var color = indexedColors.valueOf( JavaCast( "string", findColor ) );
			return color.getIndex();
		}
		catch( any exception ){
			Throw( type=exceptionType, message="Invalid Color", detail="The color provided (#arguments.colorName#) is not valid. Use getPresetColorNames() for a list of valid color names" );
		}
	}

	private any function getColor( required workbook, required string colorValue ){
		/* if colorValue is a preset name, returns the index */
		/* if colorValue is an RGB Triplet eg. "255,255,255" then the exact color object is returned for xlsx, or the nearest color's index if xls */
		var isRGB = ListLen( arguments.colorValue ) EQ 3;
		if( !isRGB )
			return getColorIndex( arguments.colorValue );
		var rgb = ListToArray( arguments.colorValue );
		if( isXmlFormat( arguments.workbook ) ){
			var rgbBytes = [
				JavaCast( "int", rgb[ 1 ] )
				,JavaCast( "int", rgb[ 2 ] )
				,JavaCast( "int", rgb[ 3 ] )
			];
			try{
				return loadClass( "org.apache.poi.xssf.usermodel.XSSFColor" ).init( JavaCast( "byte[]", rgbBytes ), JavaCast( "null", 0 ) );
			}
			//ACF doesn't handle signed java byte values the same way as Lucee: see https://www.bennadel.com/blog/2689-creating-signed-java-byte-values-using-coldfusion-numbers.htm
			catch( any exception ){
				if( !exception.message CONTAINS "cannot fit inside a byte" )
					rethrow;
				//ACF2016+ Bitwise operators can't handle >32-bit args: https://stackoverflow.com/questions/43176313/cffunction-cfargument-pass-unsigned-int32
				var javaLangInteger = CreateObject( "java", "java.lang.Integer" );
				var negativeMask = InputBaseN( ( "11111111" & "11111111" & "11111111" & "00000000" ), 2 );
				negativeMask = javaLangInteger.parseUnsignedInt( negativeMask );
				rgbBytes = [];
				for( var value in rgb ){
					if( BitMaskRead( value, 7, 1 ) )//value greater than 127
						value = BitOr( negativeMask, value );
					rgbBytes.Append( JavaCast( "byte", value ) );
				}
				return loadClass( "org.apache.poi.xssf.usermodel.XSSFColor" ).init( JavaCast( "byte[]", rgbBytes ), JavaCast( "null", 0 ) );
			}
		}
		var palette = arguments.workbook.getCustomPalette();
		var similarExistingColor = palette.findSimilarColor(
			JavaCast( "int", rgb[ 1 ] )
			,JavaCast( "int", rgb[ 2 ] )
			,JavaCast( "int", rgb[ 3 ] )
		);
		return similarExistingColor.getIndex();
	}

	public numeric function getColumnWidth( required workbook, required numeric column ){
		var columnIndex = ( arguments.column -1 );
		return ( getActiveSheet( arguments.workbook ).getColumnWidth( JavaCast( "int", columnIndex ) ) / 256 );// whole character width (of zero character)
	}

	public numeric function getColumnWidthInPixels( required workbook, required numeric column ){
		var columnIndex = ( arguments.column -1 );
		return getActiveSheet( arguments.workbook ).getColumnWidthInPixels( JavaCast( "int", columnIndex ) );
	}

	/* Override troublesome engine BIFs */

	private boolean function _isDate( required value ){
		if( !IsDate( arguments.value ) )
			return false;
		// Lucee will treat 01-23112 or 23112-01 as a date!
		if( ParseDateTime( arguments.value ).Year() > 9999 ) //ACF future limit
			return false;
		return true;
	}

	/* ACF compatibility functions */
	private array function _queryColumnArray( required query q ){
		try{
			return QueryColumnArray( arguments.q ); //Lucee
		}
		catch( any exception ){
			if( !exception.message CONTAINS "undefined" )
				rethrow;
			//ACF
			return q.getColumnNames();
		}
	}

	private query function _QueryDeleteColumn( required query q, required string columnToDelete ){
		try{
			QueryDeleteColumn( arguments.q, arguments.columnToDelete ); //Lucee/ACF2018+
			return arguments.q;
		}
		catch( any exception ){
			if( !exception.message CONTAINS "undefined" )
				rethrow;
			//ACF2016 doesn't support QueryDeleteColumn()
			var columnPosition = ListFindNoCase( arguments.q.columnList, arguments.columnToDelete );
			if( !columnPosition )
				return arguments.q;
			var columnsToKeep = ListDeleteAt( arguments.q.columnList, columnPosition );
			if( !columnsToKeep.Len() )
				return QueryNew( "" );
			return QueryExecute( "SELECT #columnsToKeep# FROM arguments.q", {}, { dbType = "query" } );
		}
	}

	private query function _queryNew( required array columnNames, required string columnTypeList, required array data ){
		//ACF QueryNew() won't accept invalid variable names in the column name list (e.g. which names including commas), hence clunky workaround:
		//NB: 'data' should not contain structs since they use the column name as key: always use array of row arrays instead
		if( !isACF )
			return QueryNew( arguments.columnNames, arguments.columnTypeList, arguments.data );
		var totalColumns = arguments.columnNames.Len();
		var tempColumnNames = [];
		for( var i=1; i LTE totalColumns; i++ )
			tempColumnNames[ i ] = "column#i#";
		var q = QueryNew( tempColumnNames.ToList(), arguments.columnTypeList, arguments.data );
		// restore the real names without ACF barfing on them
		q.setColumnNames( arguments.columnNames );
		return q;
	}

}