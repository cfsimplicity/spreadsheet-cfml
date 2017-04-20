component{

	variables.version = "1.3.0";
	variables.poiLoaderName = "_poiLoader-" & Hash( GetCurrentTemplatePath() );
	variables.javaLoaderDotPath = "javaLoader.JavaLoader";
	variables.dateFormats = {
		DATE: "yyyy-mm-dd"
		,DATETIME: "yyyy-mm-dd HH:nn:ss"
		,TIME: "hh:mm:ss"
		,TIMESTAMP: "yyyy-mm-dd hh:mm:ss"
	};
	variables.exceptionType = "cfsimplicity.lucee.spreadsheet";
	variables.isLucee5plus = ( server.KeyExists( "lucee" ) AND ( server.lucee.version.Left( 1 ) >= 5 ) );
	variables.isACF = ( server.coldfusion.productname IS "ColdFusion Server" );
	variables.engineSupportsDynamicClassLoading = isLucee5plus;
	variables.poiClassesLastLoadedVia = "Nothing loaded yet";
	variables.engineSupportsEncryption = !isACF;

	function init( struct dateFormats, string javaLoaderDotPath, boolean requiresJavaLoader ){
		if( arguments.KeyExists( "dateFormats" ) )
			overrideDefaultDateFormats( arguments.dateFormats );
		if( arguments.KeyExists( "javaLoaderDotPath" ) ) // Option to use the dot path of an existing javaloader installation to save duplication
			variables.javaLoaderDotPath = arguments.javaLoaderDotPath;
		if( arguments.KeyExists( "requiresJavaLoader" ) )
			variables.requiresJavaLoader = arguments.requiresJavaLoader;
		else
			variables.requiresJavaLoader = !engineSupportsDynamicClassLoading;
		return this;
	}

	/* Meta utilities */

	private void function overrideDefaultDateFormats( required struct formats ){
		for( var format in formats ){
			if( !variables.dateFormats.KeyExists( format ) )
				throw( type=exceptionType, message="Invalid date format key", detail="'#format#' is not a valid dateformat key. Valid keys are DATE, DATETIME, TIME and TIMESTAMP" );
			variables.dateFormats[ format ] = formats[ format ];
		}
	}

	public void function flushPoiLoader(){
		lock scope="server" timeout="10" {
			StructDelete( server, poiLoaderName );
		};
	}

	public struct function getDateFormats(){
		return dateFormats;
	}

	public struct function getEnvironment(){
		return {
			dateFormats: dateFormats
			,engine: server.coldfusion.productname & " " & ( isACF? server.coldfusion.productversion: ( server.lucee.version?: "?" ) )
			,engineSupportsDynamicClassLoading: engineSupportsDynamicClassLoading
			,engineSupportsEncryption: engineSupportsEncryption
			,javaLoaderDotPath: javaLoaderDotPath
			,poiClassesLastLoadedVia: poiClassesLastLoadedVia
			,poiLoaderName: poiLoaderName
			,requiresJavaLoader: requiresJavaLoader
		};
	}

	/* MAIN PUBLIC API */

	/* Convenenience */

	public binary function binaryFromQuery( required query data, boolean addHeaderRow=true, boldHeaderRow=true, xmlFormat=false ){
		/* Pass in a query and get a spreadsheet binary file ready to stream to the browser */
		var workbook = workbookFromQuery( argumentCollection=arguments );
		return readBinary( workbook );
	}

	public function csvToQuery(
		string csv=""
		,string filepath=""
		,boolean firstRowIsHeader=false
		,boolean trim=true
		,string delimiter
	){
		var csvIsString = csv.Len();
		var csvIsFile = filepath.Len();
		if( !csvIsString AND !csvIsFile )
			throw( type=exceptionType, message="Missing required argument", detail="Please provide either a csv string (csv), or the path of a file containing one (filepath)." );
		if( csvIsString AND csvIsFile )
			throw( type=exceptionType, message="Mutually exclusive arguments: 'csv' and 'filepath'", detail="Only one of either 'filepath' or 'csv' arguments may be provided." );
		if(	csvIsFile ){
			if( !FileExists( filepath ) )
				throw( type=exceptionType, message="Non-existant file", detail="Cannot find a file at #filepath#" );
			if( !isCsvOrTextFile( filepath ) )
				throw( type=exceptionType, message="Invalid csv file", detail="#filepath# does not appear to be a text/csv file" );
			arguments.csv = FileRead( filepath );
		}
		var format = loadPoi( "org.apache.commons.csv.CSVFormat" )[ JavaCast( "string", "RFC4180" ) ];
		format = format.withIgnoreSurroundingSpaces();//stop spaces between fields causing problems with embedded lines
		if( trim )
			csv = csv.Trim();
		if( arguments.KeyExists( "delimiter" ) )
			format = format.withDelimiter( JavaCast( "string", delimiter ) );
		var parsed = loadPoi( "org.apache.commons.csv.CSVParser" ).parse( csv, format );
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
		if( firstRowIsHeader )
			var headerRow = rows[ 1 ];
		for( var i=1; i LTE maxColumnCount; i++ ){
			if( firstRowIsHeader AND !IsNull( headerRow[ i ] ) AND headerRow[ i ].Len() ){
				columnList.Append( JavaCast( "string", headerRow[ i ] ) );
				continue;
			}
			columnList.Append( "column#i#" );
		}
		if( firstRowIsHeader )
			rows.DeleteAt( 1 );
		return _QueryNew( columnList.ToList(), "", rows );;
	}

	public void function download( required workbook, required string filename, string contentType ){
		var safeFilename = filenameSafe( filename );
		var filenameWithoutExtension = safeFilename.REReplace( "\.xlsx?$", "" );
		var extension = isXmlFormat( workbook )? "xlsx": "xls";
		arguments.filename = filenameWithoutExtension & "." & extension;
		var binary = readBinary( workbook );
		if( !arguments.KeyExists( "contentType" ) )
			arguments.contentType = isXmlFormat( workbook )? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "application/msexcel";
		downloadBinaryVariable( binary, filename, contentType );
	}

	public void function downloadFileFromQuery(
		required query data
		,required string filename
		,boolean addHeaderRow=true
		,boolean boldHeaderRow=true
		,boolean xmlFormat=false
		,string contentType
	){
		var safeFilename = filenameSafe( filename );
		var filenameWithoutExtension = safeFilename.REReplace( "\.xlsx?$","" );
		var extension = xmlFormat? "xlsx": "xls";
		arguments.filename = filenameWithoutExtension & "." & extension;
		var binary = binaryFromQuery( data,addHeaderRow,boldHeaderRow,xmlFormat );
		if( !arguments.KeyExists( "contentType" ) )
			arguments.contentType = xmlFormat? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "application/msexcel";
		downloadBinaryVariable( binary, filename, contentType );
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
		var safeFilename = filenameSafe( filename );
		var filenameWithoutExtension = safeFilename.REReplace( "\.csv$","" );
		var extension = "csv";
		arguments.filename = filenameWithoutExtension & "." & extension;
		downloadBinaryVariable( binary, filename, contentType );
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
			firstRowIsHeader: firstRowIsHeader
			,trim: trim
		};
		if( arguments.KeyExists( "csv" ) )
			conversionArgs.csv = csv;
		if( arguments.KeyExists( "filepath" ) )
			conversionArgs.filepath = filepath;
		if( arguments.KeyExists( "delimiter" ) )
			conversionArgs.delimiter = delimiter;
		var data = csvToQuery( argumentCollection=conversionArgs );
		return workbookFromQuery( data=data, addHeaderRow=firstRowIsHeader, boldHeaderRow=boldHeaderRow, xmlFormat=xmlFormat );
	}

	public any function workbookFromQuery( required query data, boolean addHeaderRow=true, boolean boldHeaderRow=true, boolean xmlFormat=false ){
		var workbook = new( xmlFormat=xmlFormat );
		if( addHeaderRow ){
			var columns = _QueryColumnArray( data );
			addRow( workbook, columns.ToList() );
			if( boldHeaderRow )
				formatRow( workbook, { bold: true }, 1 );
			addRows( workbook, data, 2, 1 );
		}
		else
			addRows( workbook, data );
		return workbook;
	}

	public void function writeFileFromQuery(
		required query data
		,required string filepath
		,boolean overwrite=false
		,boolean addHeaderRow=true
		,boldHeaderRow=true
		,xmlFormat=false
	){
		if( !xmlFormat AND ( ListLast( filepath, "." ) IS "xlsx" ) )
			arguments.xmlFormat = true;
		var workbook = workbookFromQuery( data, addHeaderRow, boldHeaderRow, xmlFormat );
		if( xmlFormat AND ( ListLast( filepath, "." ) IS "xls" ) )
			arguments.filepath &= "x";// force to .xlsx
		write( workbook=workbook, filepath=filepath, overwrite=overwrite );
	}

	/* End convenience methods */

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
		var rowNum = ( arguments.KeyExists( "startRow" ) AND startRow )? startRow-1: 0;
		var cellNum = 0;
		var lastCellNum = 0;
		var cellValue = 0;
		if( arguments.KeyExists( "startColumn" ) )
			cellNum = ( startColumn -1 );
		else {
			row = getActiveSheet( workbook ).getRow( rowNum );
			/* if this row exists, find the next empty cell number. note: getLastCellNum() returns the cell index PLUS ONE or -1 if not found */
			if( !IsNull( row ) AND row.getLastCellNum() GT 0 )
				cellNum = row.getLastCellNum();
			else
				cellNum = 0;
		}
		var columnNumber = ( cellNum +1 );
		var columnData = ListToArray( data, delimiter );
		for( var cellValue in columnData ){
			/* if rowNum is greater than the last row of the sheet, need to create a new row  */
			if( rowNum GT getActiveSheet( workbook ).getLastRowNum() OR IsNull( getActiveSheet( workbook ).getRow( rowNum ) ) )
				row = createRow( workbook, rowNum );
			else
				row = getActiveSheet( workbook ).getRow( rowNum );
			/* POI doesn't have any 'shift column' functionality akin to shiftRows() so inserts get interesting */
			/* ** Note: row.getLastCellNum() returns the cell index PLUS ONE or -1 if not found */
			if( insert AND ( cellNum LT row.getLastCellNum() ) ){
				/*  need to get the last populated column number in the row, figure out which cells are impacted, and shift the impacted cells to the right to make room for the new data */
				lastCellNum = row.getLastCellNum();
				for( var i = lastCellNum; i EQ cellNum; i-- ){
					oldCell = row.getCell( JavaCast( "int", i-1 ) );
					if( !IsNull( oldCell ) ){
						cell = createCell( row, i );
						cell.setCellStyle( oldCell.getCellStyle() );
						var cellValue = getCellValueAsType( workbook, oldCell );
						setCellValueAsType( workbook, oldCell, cellValue );
						cell.setCellComment( oldCell.getCellComment() );
					}
				}
			}
			cell = createCell( row,cellNum );
			setCellValueAsType( workbook, cell, cellValue );
			rowNum++;
		}
		if( autoSize )
			autoSizeColumn( workbook, columnNumber );
	}

	public void function addFreezePane(
		required workbook
		,required numeric freezeColumn
		,required numeric freezeRow
		,numeric leftmostColumn //left column visible in right pane
		,numeric topRow //top row visible in bottom pane
	){
		if( arguments.KeyExists( "leftmostColumn" ) AND !arguments.KeyExists( "topRow" ) )
			arguments.topRow = freezeRow;
		if( arguments.KeyExists( "topRow" ) AND !arguments.KeyExists( "leftmostColumn" ) )
			arguments.leftmostColumn = freezeColumn;
		/* createFreezePane() operates on the logical row/column numbers as opposed to physical, so no need for n-1 stuff here */
		if( !arguments.KeyExists( "leftmostColumn" ) ){
			getActiveSheet( workbook ).createFreezePane( JavaCast( "int", freezeColumn ),JavaCast( "int", freezeRow ) );
			return;
		}
		// POI lets you specify an active pane if you use createSplitPane() here
		getActiveSheet( workbook ).createFreezePane(
			JavaCast( "int", freezeColumn )
			,JavaCast( "int", freezeRow )
			,JavaCast( "int", leftmostColumn )
			,JavaCast( "int", topRow )
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
			throw( type=exceptionType, message="Invalid argument combination", detail="You must provide either a file path or an image object" );
		if( arguments.KeyExists( "imageData" ) AND !arguments.KeyExists( "imageType" ) )
			throw( type=exceptionType, message="Invalid argument combination", detail="If you specify an image object, you must also provide the imageType argument" );
		var numberOfAnchorElements = ListLen( anchor );
		if( ( numberOfAnchorElements NEQ 4 ) AND ( numberOfAnchorElements NEQ 8 ) )
			throw( type=exceptionType, message="Invalid anchor argument", detail="The anchor argument must be a comma-delimited list of integers with either 4 or 8 elements" );
		//we'll need the image type int in all cases
		if( arguments.KeyExists( "filepath" ) ){
			if( !FileExists( filepath ) )
				throw( type=exceptionType, message="Non-existent file", detail="The specified file does not exist." );
			try{
				arguments.imageType = ListLast( FileGetMimeType( filepath ), "/" );
			}
			catch( any exception ){
				throw( type=exceptionType, message="Could Not Determine Image Type", detail="An image type could not be determined from the filepath provided" );
			}
		}
		else if( !arguments.KeyExists( "imageType" ) )
			throw( type=exceptionType,message="Could Not Determine Image Type",detail="An image type could not be determined from the filepath or imagetype provided" );
		arguments.imageType	=	imageType.UCase();
		switch( imageType ){
			case "DIB": case "EMF": case "JPEG": case "PICT": case "PNG": case "WMF":
				var imageTypeIndex = workbook[ "PICTURE_TYPE_" & imageType ];
			break;
			case "JPG":
				var imageTypeIndex = workbook.PICTURE_TYPE_JPEG;
			break;
			default:
				throw( type=exceptionType, message="Invalid Image Type", detail="Valid image types are DIB, EMF, JPG, JPEG, PICT, PNG, and WMF" );
		}
		if( arguments.KeyExists( "filepath" ) ){
			try{
				var inputStream=CreateObject( "java","java.io.FileInputStream" ).init( JavaCast("string",filepath ) );
				var ioUtils=loadPoi( "org.apache.poi.util.IOUtils" );
				var bytes=ioUtils.toByteArray( inputStream );
			}
			finally{
				if( local.KeyExists( "inputStream" ) )
					inputStream.close();
			}
		}
		else
			var bytes = ToBinary( imageData );
		var imageIndex = workbook.addPicture( bytes, JavaCast( "int", imageTypeIndex ) );
		var clientAnchorClass = isXmlFormat( workbook )
				? "org.apache.poi.xssf.usermodel.XSSFClientAnchor"
				: "org.apache.poi.hssf.usermodel.HSSFClientAnchor";
		var theAnchor = loadPoi( clientAnchorClass ).init();
		if( numberOfAnchorElements EQ 4 ){
			theAnchor.setRow1( JavaCast( "int", ListFirst( anchor )-1 ) );
			theAnchor.setCol1( JavaCast( "int", ListGetAt( anchor, 2 )-1 ) );
			theAnchor.setRow2( JavaCast( "int", ListGetAt( anchor, 3 )-1 ) );
			theAnchor.setCol2( JavaCast( "int", ListLast( anchor )-1 ) );
		} else if( numberOfAnchorElements EQ 8 ){
			theAnchor.setDx1( JavaCast( "int", ListFirst( anchor ) ) );
			theAnchor.setDy1( JavaCast( "int", ListGetAt( anchor,2 ) ) );
			theAnchor.setDx2( JavaCast( "int", ListGetAt( anchor,3 ) ) );
			theAnchor.setDy2( JavaCast( "int", ListGetAt( anchor,4 ) ) );
			theAnchor.setRow1( JavaCast( "int", ListGetAt( anchor,5 )-1 ) );
			theAnchor.setCol1( JavaCast( "int", ListGetAt( anchor,6 )-1 ) );
			theAnchor.setRow2( JavaCast( "int", ListGetAt( anchor,7 )-1 ) );
			theAnchor.setCol2( JavaCast( "int", ListLast( anchor )-1 ) );
		}
		/* TODO: need to look into createDrawingPatriarch() vs. getDrawingPatriarch() since create will kill any existing images. getDrawingPatriarch() throws  a null pointer exception when an attempt is made to add a second image to the spreadsheet  */
		var drawingPatriarch = getActiveSheet( workbook ).createDrawingPatriarch();
		var picture = drawingPatriarch.createPicture( theAnchor, imageIndex );
		/* Disabling this for now--maybe let people pass in a boolean indicating whether or not they want the image resized?
		 if this is a png or jpg, resize the picture to its original size (this doesn't work for formats other than jpg and png)
			<cfif imgTypeIndex eq getWorkbook().PICTURE_TYPE_JPEG or imgTypeIndex eq getWorkbook().PICTURE_TYPE_PNG>
				<cfset picture.resize() />
			</cfif>
		*/
	}

	public void function addInfo( required workbook,required struct info ){
		/* Valid struct keys are author, category, lastauthor, comments, keywords, manager, company, subject, title */
		if( isBinaryFormat( workbook ) )
			addInfoBinary( workbook,info );
		else
			addInfoXml( workbook,info );
	}

	public void function addRow(
		required workbook
		,required string data /* Delimited list of data */
		,numeric row
		,numeric column=1
		,boolean insert=true
		,string delimiter=","
		,boolean handleEmbeddedCommas=true /* When true, values enclosed in single quotes are treated as a single element like in ACF. Only applies when the delimiter is a comma. */
		,boolean autoSizeColumns=false
	){
		if( arguments.KeyExists( "row" ) AND ( row LTE 0 ) )
			throw( type=exceptionType, message="Invalid row value", detail="The value for row must be greater than or equal to 1." );
		if( arguments.KeyExists( "column" ) AND ( column LTE 0 ) )
			throw( type=exceptionType, message="Invalid column value", detail="The value for column must be greater than or equal to 1." );
		if( !insert AND !arguments.KeyExists( "row") )
			throw( type=exceptionType, message="Missing row value", detail="To replace a row using 'insert', please specify the row to replace." );
		var lastRow = getNextEmptyRow( workbook );
		//If the requested row already exists...
		if( arguments.KeyExists( "row" ) AND ( row LTE lastRow ) ){
			if( arguments.insert )
				shiftRows( workbook, row, lastRow, 1 );//shift the existing rows down (by one row)
			else
				deleteRow( workbook, row );//otherwise, clear the entire row
		}
		var theRow = arguments.KeyExists( "row" )? createRow( workbook,arguments.row-1 ): createRow( workbook );
		var rowValues = parseRowData( data, delimiter, handleEmbeddedCommas );
		var cellIndex = column-1;
		for( var cellValue in rowValues ){
			var cell = createCell( theRow, cellIndex );
			setCellValueAsType( workbook, cell, Trim( cellValue ) );
			if( autoSizeColumns )
				autoSizeColumn( workbook, column );
			cellIndex++;
		}
	}

	public void function addRows(
		required workbook
		,required query data
		,numeric row
		,numeric column=1
		,boolean insert=true
		,boolean autoSizeColumns=false
		,boolean includeQueryColumnNames=false
	){
		var lastRow = getNextEmptyRow( workbook );
		var insertAtRowIndex = arguments.keyExists( "row" )? row-1: getNextEmptyRow( workbook );
		if( arguments.KeyExists( "row" ) AND ( row LTE lastRow ) AND insert )
			shiftRows( workbook,row, lastRow, data.recordCount );
		var currentRowIndex = insertAtRowIndex;
		var queryColumns = getQueryColumnFormats( workbook, data );
		var dateUtil = getDateUtil();
		for( var dataRow in data ){
			var newRow=createRow( workbook, currentRowIndex, false );
			var cellIndex = ( column -1 );
   		/* populate all columns in the row */
   		for( var queryColumn in queryColumns ){
   			var cell = createCell( newRow, cellIndex, false );
				var value = dataRow[ queryColumn.name ];
				queryColumn.index = cellIndex;
				/* Cast the values to the correct type  */
				switch( queryColumn.cellDataType ){
					case "DOUBLE":
						setCellValueAsType( workbook, cell, value, "numeric" );
						break;
					case "DATE":
					case "TIME":
						setCellValueAsType( workbook, cell, value, "date" );
						break;
					case "BOOLEAN":
						setCellValueAsType( workbook, cell, value, "boolean" );
						break;
					default:
						if( IsSimpleValue( value ) AND !Len( value ) ) //NB don't use member function: won't work if numeric
							setCellValueAsType( workbook, cell, value, "blank" );
						else
							setCellValueAsType( workbook, cell, value, "string" );
				}
				cellIndex++;
   		}
   		currentRowIndex++;
		}
		if( autoSizeColumns ){
			var numberOfColumns = queryColumns.Len();
			var thisColumn = column;
			for( var i = thisColumn; i LTE numberOfColumns; i++ ){
				autoSizeColumn( workbook, thisColumn );
				thisColumn++;
			}
		}
		if( includeQueryColumnNames ){
			var columnNames = _QueryColumnArray( data );
			var delimiter = "|";
			var columnNamesList = columnNames.ToList( delimiter );
			addRow( workbook=workbook, data=columnNamesList, row=insertAtRowIndex+1, delimiter=delimiter );
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
		var activeSheet = getActiveSheet( workbook );
		arguments.activePane = activeSheet[ "PANE_#activePane#" ];
		activeSheet.createSplitPane(
			JavaCast( "int", xSplitPosition )
			,JavaCast( "int", ySplitPosition )
			,JavaCast( "int", leftmostColumn )
			,JavaCast( "int", topRow )
			,JavaCast( "int", activePane )
		);
	}

	public void function autoSizeColumn( required workbook, required numeric column, boolean useMergedCells=false ){
		if( column LTE 0 )
			throw( type=exceptionType, message="Invalid column value", detail="The value for column must be greater than or equal to 1." );
		/* Adjusts the width of the specified column to fit the contents. For performance reasons, this should normally be called only once per column. */
		var columnIndex = column-1;
		getActiveSheet( workbook ).autoSizeColumn( columnIndex, useMergedCells );
	}

	public void function clearCell( required workbook, required numeric row, required numeric column ){
		/* Clears the specified cell of all styles and values */
		var defaultStyle = workbook.getCellStyleAt( JavaCast( "short", 0 ) );
		var rowIndex = ( row -1 );
		var rowObject = getActiveSheet( workbook ).getRow( JavaCast( "int", rowIndex ) );
		if( IsNull( rowObject ) )
			return;
		var columnIndex = ( column -1 );
		var cell = rowObject.getCell( JavaCast( "int", columnIndex ) );
		if( IsNull( cell ) )
			return;
		cell.setCellStyle( defaultStyle );
		cell.setCellType( cell.CELL_TYPE_BLANK );
	}

	public void function clearCellRange(
		required workbook
		,required numeric startRow
		,required numeric startColumn
		,required numeric endRow
		,required numeric endColumn
	){
		/* Clears the specified cell range of all styles and values */
		for( var rowNumber = startRow; rowNumber LTE endRow; rowNumber++ ){
			for( var columnNumber = startColumn; columnNumber LTE endColumn; columnNumber++ ){
				clearCell( workbook, rowNumber, columnNumber );
			}
		}
	}

	public void function createSheet( required workbook, string sheetName, overwrite=false ){
		if( arguments.KeyExists( "sheetName" ) )
			validateSheetName( sheetName );
		else
			arguments.sheetName = generateUniqueSheetName( workbook );
		if( !sheetExists( workbook=workbook, sheetName=sheetName ) ){
			workbook.createSheet( JavaCast( "String", sheetName ) );
			return;
		}
		/* sheet already exists with that name */
		if( !overwrite )
			throw( type=exceptionType, message="Sheet name already exists", detail="A sheet with the name '#sheetName#' already exists in this workbook" );
		/* OK to replace the existing */
		var sheetIndexToReplace = workbook.getSheetIndex( JavaCast( "string", sheetName ) );
		deleteSheetAtIndex( workbook, sheetIndexToReplace );
		var newSheet = workbook.createSheet( JavaCast( "String", sheetName ) );
		var moveToIndex = sheetIndexToReplace;
		moveSheet( workbook, sheetName, moveToIndex );
	}

	public void function deleteColumn( required workbook,required numeric column ){
		if( column LTE 0 )
			throw( type=exceptionType, message="Invalid column value", detail="The value for column must be greater than or equal to 1." );
			/* POI doesn't have remove column functionality, so iterate over all the rows and remove the column indicated */
		var rowIterator = getActiveSheet( workbook ).rowIterator();
		while( rowIterator.hasNext() ){
			var row = rowIterator.next();
			var cell = row.getCell( JavaCast( "int", ( column -1 ) ) );
			if( IsNull( cell ) )
				continue;
			row.removeCell( cell );
		}
	}

	public void function deleteColumns( required workbook, required string range ){
		/* Validate and extract the ranges. Range is a comma-delimited list of ranges, and each value can be either a single number or a range of numbers with a hyphen. */
		var allRanges = extractRanges( range );
		for( var thisRange in allRanges ){
			if( thisRange.startAt EQ thisRange.endAt ){
				/* Just one row */
				deleteColumn( workbook, thisRange.startAt );
				continue;
			}
			for( var columnNumber = thisRange.startAt; columnNumber LTE thisRange.endAt; columnNumber++ )
				deleteColumn( workbook,columnNumber );
		}
	}

	public void function deleteRow( required workbook, required numeric row ){
		/* Deletes the data from a row. Does not physically delete the row. */
		if( row LTE 0 )
			throw( type=exceptionType, message="Invalid row value", detail="The value for row must be greater than or equal to 1." );
		var rowToDelete = row-1;
		if( rowToDelete GTE getFirstRowNum( workbook ) AND rowToDelete LTE getLastRowNum( workbook ) ) //If this is a valid row, remove it
			getActiveSheet( workbook ).removeRow( getActiveSheet( workbook ).getRow( JavaCast( "int", rowToDelete ) ) );
	}

	public void function deleteRows( required workbook, required string range ){
		/* Validate and extract the ranges. Range is a comma-delimited list of ranges, and each value can be either a single number or a range of numbers with a hyphen. */
		var allRanges = extractRanges( range );
		for( var thisRange in allRanges ){
			if( thisRange.startAt EQ thisRange.endAt ){
				/* Just one row */
				deleteRow( workbook, thisRange.startAt );
				continue;
			}
			for( var rowNumber = thisRange.startAt; rowNumber LTE thisRange.endAt; rowNumber++ )
				deleteRow( workbook, rowNumber );
		}
	}

	public void function formatCell( required workbook, required struct format, required numeric row, required numeric column, any cellStyle ){
		var cell = initializeCell( workbook, row, column );
		var style = arguments.cellStyle?: buildCellStyle( workbook, format );
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
		var style = arguments.cellStyle?: buildCellStyle( workbook,format );
		for( var rowNumber = startRow; rowNumber LTE endRow; rowNumber++ ){
			for( var columnNumber = startColumn; columnNumber LTE endColumn; columnNumber++ )
				formatCell( workbook, format, rowNumber, columnNumber, style );
		}
	}

	public void function formatColumn( required workbook, required struct format, required numeric column, any cellStyle ){
		if( column LT 1 )
			throw( type=exceptionType, message="Invalid column value", detail="The column value must be greater than 0" );
		var style = arguments.cellStyle?: buildCellStyle( workbook,format );
		var rowIterator = getActiveSheet( workbook ).rowIterator();
		var columnNumber = column;
		while( rowIterator.hasNext() ){
			var rowNumber = rowIterator.next().getRowNum() + 1;
			formatCell( workbook, format, rowNumber, columnNumber, style );
		}
	}

	public void function formatColumns( required workbook, required struct format, required string range, any cellStyle ){
		/* Validate and extract the ranges. Range is a comma-delimited list of ranges, and each value can be either a single number or a range of numbers with a hyphen. */
		var allRanges = extractRanges( range );
		var style = arguments.cellStyle?: buildCellStyle( workbook, format );
		for( var thisRange in allRanges ){
			if( thisRange.startAt EQ thisRange.endAt ){
				/* Just one column */
				formatColumn( workbook, format, thisRange.startAt, style );
				continue;
			}
			for( var columnNumber = thisRange.startAt; columnNumber LTE thisRange.endAt; columnNumber++ )
				formatColumn( workbook, format, columnNumber, style );
		}
	}

	public void function formatRow( required workbook, required struct format, required numeric row, any cellStyle ){
		var rowIndex = row-1;
		var theRow = getActiveSheet( workbook ).getRow( rowIndex );
		if( IsNull( theRow ) )
			return;
		var style = arguments.cellStyle?: buildCellStyle( workbook, format );
		var cellIterator = theRow.cellIterator();
		while( cellIterator.hasNext() )
			formatCell( workbook, format, row, ( cellIterator.next().getColumnIndex() +1 ), style );
	}

	public void function formatRows( required workbook, required struct format, required string range, any cellStyle ){
		/* Validate and extract the ranges. Range is a comma-delimited list of ranges, and each value can be either a single number or a range of numbers with a hyphen. */
		var allRanges = extractRanges( range );
		var style = arguments.cellStyle?: buildCellStyle( workbook,format );
		for( var thisRange in allRanges ){
			if( thisRange.startAt EQ thisRange.endAt ){
				/* Just one row */
				formatRow( workbook, format, thisRange.startAt, style );
				continue;
			}
			for( var rowNumber = thisRange.startAt; rowNumber LTE thisRange.endAt; rowNumber++ )
				formatRow( workbook, format, rowNumber, style );
		}
	}

	public any function getCellComment( required workbook, numeric row, numeric column ){
		if( arguments.keyExists( "row" ) AND !arguments.KeyExists( "column" ) )
			throw( type=exceptionType, message="Invalid argument combination", detail="If you specify the row you must also specify the column" );
		if( arguments.keyExists( "column" ) AND !arguments.KeyExists( "row" ) )
			throw( type=exceptionType, message="Invalid argument combination", detail="If you specify the column you must also specify the row" );
		if( arguments.KeyExists( "row" ) ){
			var cell = getCellAt( workbook, row, column );
			var commentObject = cell.getCellComment();
			if( !IsNull( commentObject ) ){
				return {
					author: commentObject.getAuthor()
					,comment: commentObject.getString().getString()
					,column: column
					,row: row
				};
			}
			return {};
		}
		/* TODO: Look into checking all sheets in the workbook */
		/* row and column weren't provided so loop over the whole sheet and return all the comments as an array of structs */
		var result = [];
		var rowIterator = getActiveSheet( workbook ).rowIterator();
		while( rowIterator.hasNext() ){
			var cellIterator = rowIterator.next().cellIterator();
			while( cellIterator.hasNext() ){
				var commentObject = cellIterator.next().getCellComment();
				if( !IsNull( commentObject ) ){
					var comment = {
						author: commentObject.getAuthor()
						,comment: commentObject.getString().getString()
						,column: column
						,row: row
					};
					comments.Append( comment );
				}
			}
		}
		return comments;
	}

	public any function getCellFormula( required workbook, numeric row, numeric column ){
		if( arguments.KeyExists( "row" ) AND arguments.KeyExists( "column" ) ){
			if( cellExists( workbook, row, column ) ){
				var cell = getCellAt( workbook, row,column );
				if( cell.getCellType() IS cell.CELL_TYPE_FORMULA )
					return cell.getCellFormula();
				return "";
			}
		}
		//no row and column provided so return an array of structs containing formulas for the entire sheet
		var rowIterator = getActiveSheet( workbook ).rowIterator();
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
					formulaStruct.formula=cell.getCellFormula();
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
		if( !cellExists( workbook, row, column ) )
			return "";
		var rowIndex = ( row-1 );
		var columnIndex = ( column-1 );
		var rowObject = getActiveSheet( workbook ).getRow( JavaCast( "int", rowIndex ) );
		var cell = rowObject.getCell( JavaCast( "int", columnIndex ) );
		return cell.getCellTypeEnum().toString();
	}

	public any function getCellValue( required workbook, required numeric row, required numeric column ){
		if( !cellExists( workbook, row, column ) )
			return "";
		var rowIndex = ( row-1 );
		var columnIndex = ( column-1 );
		var rowObject = getActiveSheet( workbook ).getRow( JavaCast( "int", rowIndex ) );
		var cell = rowObject.getCell( JavaCast( "int", columnIndex ) );
		var formatter = getFormatter();
		if( cell.getCellType() EQ cell.CELL_TYPE_FORMULA ){
			var formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
			return formatter.formatCellValue( cell, formulaEvaluator );
		}
		return formatter.formatCellValue( cell );
	}

	public numeric function getColumnCount( required workbook, sheetNameOrNumber ){
		if( arguments.KeyExists( "sheetNameOrNumber" ) ){
			if( IsValid( "integer", sheetNameOrNumber ) AND IsNumeric( sheetNameOrNumber ) ){
				var sheetNumber = sheetNameOrNumber;
				validateSheetNumber( workbook, sheetNumber );
			} else {
				var sheetName = sheetNameOrNumber;
				validateSheetExistsWithName( workbook, sheetName );
				var sheetNumber = workbook.getSheetIndex( JavaCast( "string", sheetName ) ) + 1;
			}
			setActiveSheetNumber( workbook, sheetNumber );
		}
		var sheet = getActiveSheet( workbook );
		var rowIterator = sheet.rowIterator();
		var result = 0;
		while( rowIterator.hasNext() ){
			var row = rowIterator.next();
			result = Max( result, row.getLastCellNum() );
		}
		return result;
	}

	public void function hideColumn( required workbook, required numeric column ){
		toggleColumnHidden( workbook, column, true );
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
		return workbook.getClass().getCanonicalName() IS "org.apache.poi.hssf.usermodel.HSSFWorkbook";
	}

	public boolean function isSpreadsheetFile( required string path ){
		if( !FileExists( path ) )
			throw( type=exceptionType, message="Non-existent file", detail="Cannot find the file #path#." );
		try{
			var workbook = workbookFromFile( path );
		}
		catch( cfsimplicity.lucee.spreadsheet.invalidFile exception ){
			return false;
		}
		return true;
	}

	public boolean function isSpreadsheetObject( required object ){
		return isBinaryFormat( object ) OR isXmlFormat( object );
	}

	public boolean function isXmlFormat( required workbook ){
		return workbook.getClass().getCanonicalName() IS "org.apache.poi.xssf.usermodel.XSSFWorkbook";
	}

	public void function mergeCells(
		required workbook
		,required numeric startRow
		,required numeric endRow
		,required numeric startColumn
		,required numeric endColumn
		,boolean emptyInvisibleCells=false
	){
		if( startRow LT 1 OR startRow GT endRow )
			throw( type=exceptionType, message="Invalid startRow or endRow", detail="Row values must be greater than 0 and the startRow cannot be greater than the endRow." );
		if( startColumn LT 1 OR startColumn GT endColumn )
			throw( type=exceptionType, message="Invalid startColumn or endColumn", detail="Column values must be greater than 0 and the startColumn cannot be greater than the endColumn." );
		var cellRangeAddress = loadPoi( "org.apache.poi.ss.util.CellRangeAddress" ).init(
			JavaCast( "int", ( startRow - 1 ) )
			,JavaCast( "int", ( endRow - 1 ) )
			,JavaCast( "int", ( startColumn - 1 ) )
			,JavaCast( "int", ( endColumn - 1 ) )
		);
		getActiveSheet( workbook ).addMergedRegion( cellRangeAddress );
		if( !emptyInvisibleCells )
			return;
		// stash the value to retain
		var visibleValue = getCellValue( workbook, startRow, startColumn );
		//empty all cells in the merged region
		setCellRangeValue( workbook, "", startRow, endRow, startColumn, endColumn );
		//restore the stashed value
		setCellValue( workbook, visibleValue, startRow, startColumn );
	}

	public any function new( string sheetName="Sheet1", boolean xmlformat=false ){
		var workbook = createWorkBook( sheetName,xmlFormat );
		createSheet( workbook,sheetName, xmlformat );
		setActiveSheet( workbook, sheetName );
		return workbook;
	}

	public any function newXls( string sheetName="Sheet1" ){
		return new( sheetName=sheetName, xmlFormat=false );
	}

	public any function newXlsx( string sheetName="Sheet1" ){
		return new( sheetName=sheetName, xmlFormat=true );
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
			throw( type=exceptionType, message="Invalid argument 'query'.", detail="Just use format='query' to return a query object" );
		if( arguments.KeyExists( "format" ) AND !ListFindNoCase( "query,html,csv", format ) )
			throw( type=exceptionType, message="Invalid format", detail="Supported formats are: 'query', 'html' and 'csv'" );
		if( arguments.KeyExists( "sheetName" ) AND arguments.KeyExists( "sheetNumber" ) )
			throw( type=exceptionType, message="Cannot provide both sheetNumber and sheetName arguments", detail="Only one of either 'sheetNumber' or 'sheetName' arguments may be provided." );
		if( !FileExists( src ) )
			throw( type=exceptionType, message="Non-existent file", detail="Cannot find the file #src#." );
		var passwordProtected = ( arguments.KeyExists( "password") AND !password.Trim().IsEmpty() );
		if( passwordProtected AND !engineSupportsEncryption )
			throw( type=exceptionType, message="Reading password protected files is not supported for Adobe ColdFusion", detail="Reading password protected files currently only works in Lucee, not ColdFusion" );
		var workbook = passwordProtected? decryptFile( src, password ): workbookFromFile( src );
		if( arguments.KeyExists( "sheetName" ) )
			setActiveSheet( workbook=workbook, sheetName=sheetName );
		if( !arguments.keyExists( "format" ) )
			return workbook;
		var args = {
			workbook: workbook
		};
		if( arguments.KeyExists( "sheetName" ) )
			args.sheetName = sheetName;
		if( arguments.KeyExists( "sheetNumber" ) )
			args.sheetNumber = sheetNumber;
		if( arguments.KeyExists( "headerRow" ) ){
			args.headerRow = headerRow;
			args.includeHeaderRow = includeHeaderRow;
		}
		if( arguments.KeyExists( "rows" ) )
			args.rows = rows;
		if( arguments.KeyExists( "columns" ) )
			args.columns = columns;
		if( arguments.KeyExists( "columnNames" ) )
			args.columnNames = columnNames;
		args.includeBlankRows = includeBlankRows;
		args.fillMergedCellsWithVisibleValue = fillMergedCellsWithVisibleValue;
		args.includeHiddenColumns = includeHiddenColumns;
		args.includeRichTextFormatting = includeRichTextFormatting;
		var generatedQuery = sheetToQuery( argumentCollection=args );
		if( format IS "query" )
			return generatedQuery;
		var args = {
			query: generatedQuery
		};
		if( arguments.KeyExists( "headerRow" ) ){
			args.headerRow = headerRow;
			args.includeHeaderRow = includeHeaderRow;
		}
		switch( format ){
			case "csv": return queryToCsv( argumentCollection=args );
			case "html": return queryToHtml( argumentCollection=args );
		}
	}

	public binary function readBinary( required workbook ){
		var baos = CreateObject( "Java", "org.apache.commons.io.output.ByteArrayOutputStream" ).init();
		workbook.write( baos );
		baos.flush();
		return baos.toByteArray();
	}

	public void function removeSheet( required workbook, required string sheetName ){
		validateSheetName( sheetName );
		validateSheetExistsWithName( workbook, sheetName );
		arguments.sheetNumber = ( workbook.getSheetIndex( sheetName ) +1 );
		var sheetIndex = ( sheetNumber -1 );
		deleteSheetAtIndex( workbook, sheetIndex );
	}

	public void function removeSheetNumber( required workbook, required numeric sheetNumber ){
		validateSheetNumber( workbook, sheetNumber );
		var sheetIndex = sheetNumber-1;
		deleteSheetAtIndex( workbook, sheetIndex );
	}

	public void function renameSheet( required workbook, required string sheetName, required numeric sheetNumber ){
		validateSheetName( sheetName );
		validateSheetNumber( workbook, sheetNumber );
		var sheetIndex = sheetNumber-1;
		var foundAt = workbook.getSheetIndex( JavaCast( "string", sheetName ) );
		if( ( foundAt GT 0 ) AND ( foundAt NEQ sheetIndex ) )
			throw( type=exceptionType, message="Invalid Sheet Name [#sheetName#]", detail="The workbook already contains a sheet named [#sheetName#]. Sheet names must be unique" );
		workbook.setSheetName( JavaCast( "int", sheetIndex ), JavaCast( "string", sheetName ) );
	}

	public void function setActiveSheet( required workbook, string sheetName, numeric sheetNumber ){
		validateSheetNameOrNumberWasProvided( argumentCollection=arguments );
		if( arguments.KeyExists( "sheetName" ) ){
			validateSheetExistsWithName( workbook, sheetName );
			sheetNumber = ( workbook.getSheetIndex( JavaCast( "string", sheetName ) ) + 1 );
		}
		validateSheetNumber( workbook,sheetNumber );
		workbook.setActiveSheet( JavaCast( "int", (sheetNumber - 1 ) ) );
	}

	public void function setActiveSheetNumber( required workbook, numeric sheetNumber ){
		setActiveSheet( argumentCollection=arguments );
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
		var drawingPatriarch = getActiveSheet( workbook ).createDrawingPatriarch();
		var commentString = loadPoi( "org.apache.poi.hssf.usermodel.HSSFRichTextString" ).init( JavaCast( "string", comment.comment ) );
		var javaColorRGB = 0;
		if( comment.KeyExists( "anchor" ) )
			var clientAnchor = loadPoi( "org.apache.poi.hssf.usermodel.HSSFClientAnchor" ).init(
				JavaCast( "int", 0 )
				,JavaCast( "int", 0 )
				,JavaCast( "int", 0 )
				,JavaCast( "int", 0 )
				,JavaCast( "int", ListGetAt( comment.anchor, 1 ) )
				,JavaCast( "int", ListGetAt( comment.anchor, 2 ) )
				,JavaCast( "int", ListGetAt( comment.anchor, 3 ) )
				,JavaCast( "int", ListGetAt( comment.anchor, 4 ) )
			);
		else
			var clientAnchor = loadPoi( "org.apache.poi.hssf.usermodel.HSSFClientAnchor" ).init(
				JavaCast( "int", 0 )
				,JavaCast( "int", 0 )
				,JavaCast( "int", 0 )
				,JavaCast( "int", 0 )
				,JavaCast( "int", column )
				,JavaCast( "int", row )
				,JavaCast( "int", ( column +2 ) )
				,JavaCast( "int", ( row +2 ) )
			);
		var commentObject = drawingPatriarch.createComment( clientAnchor );
		if( comment.KeyExists( "author" ) )
			commentObject.setAuthor( JavaCast( "string", comment.author ) );
		/* If we're going to do anything font related, need to create a font. Didn't really want to create it above since it might not be needed.  */
		if( comment.KeyExists( "bold" )
				OR comment.KeyExists( "color" )
				OR comment.KeyExists( "font" )
				OR comment.KeyExists( "italic" )
				OR comment.KeyExists( "size" )
				OR comment.KeyExists( "strikeout" )
				OR comment.KeyExists( "underline" )
		){
			var font = workbook.createFont();
			if( comment.KeyExists( "bold" ) )
				font.setBold( JavaCast( "boolean", comment.bold ) );
			if( comment.KeyExists( "color" ) )
				font.setColor( getColor( workbook, comment.color ) );
			if( comment.KeyExists( "font" ) )
				font.setFontName( JavaCast( "string", comment.font ) );
			if( comment.KeyExists( "italic" ) )
				font.setItalic( JavaCast( "string", comment.italic ) );
			if( comment.KeyExists( "size" ) )
				font.setFontHeightInPoints( JavaCast( "int", comment.size ) );
			if( comment.KeyExists( "strikeout" ) )
				font.setStrikeout( JavaCast( "boolean", comment.strikeout ) );
			if( comment.KeyExists( "underline" ) )
				font.setUnderline( JavaCast( "boolean", comment.underline ) );
			commentString.applyFont( font );
		}
		if( comment.KeyExists( "fillColor" ) ){
			javaColorRGB = getJavaColorRGB( comment.fillColor );
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
		if( comment.KeyExists( "horizontalAlignment" ) ){
			if( comment.horizontalAlignment.UCase() IS "CENTER" )
				comment.horizontalAlignment="CENTERED";
			if( comment.horizontalAlignment.UCase() IS "JUSTIFY" )
				comment.horizontalAlignment="JUSTIFIED";
			commentObject.setHorizontalAlignment( JavaCast( "int", commentObject[ "HORIZONTAL_ALIGNMENT_" & comment.horizontalalignment.UCase() ] ) );
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
		if( comment.KeyExists( "lineStyle" ) )
		 	commentObject.setLineStyle( JavaCast( "int", commentObject[ "LINESTYLE_" & comment.lineStyle.UCase() ] ) );
		if( comment.KeyExists( "lineStyleColor" ) ){
			javaColorRGB = getJavaColorRGB( comment.lineStyleColor );
			commentObject.setLineStyleColor(
				JavaCast( "int", javaColorRGB.red )
				,JavaCast( "int", javaColorRGB.green )
				,JavaCast( "int", javaColorRGB.blue )
			);
		}
		/* Vertical alignment can be top, center, bottom, justify, and distributed. Note that center and justify are DIFFERENT than the constants for horizontal alignment, which are CENTERED and JUSTIFIED. */
		if( comment.KeyExists( "verticalAlignment" ) )
			commentObject.setVerticalAlignment( JavaCast( "int", commentObject[ "VERTICAL_ALIGNMENT_" & comment.verticalAlignment.UCase() ] ) );
		if( comment.KeyExists( "visible" ) )
			commentObject.setVisible( JavaCast( "boolean", comment.visible ) );//doesn't seem to work
		commentObject.setString( commentString );
		var cell = initializeCell( workbook, row, column );
		cell.setCellComment( commentObject );
	}

	public void function setCellFormula( required workbook, required string formula, required numeric row, required numeric column ){
		//Automatically create the cell if it does not exist, instead of throwing an error
		var cell = initializeCell( workbook, row, column );
		cell.setCellFormula( JavaCast( "string", formula ) );
	}

	public void function setCellValue( required workbook, required value, required numeric row, required numeric column, string type ){
		//Automatically create the cell if it does not exist, instead of throwing an error
		var args = {
			workbook: workbook
			,cell: initializeCell( workbook, row, column )
			,value: value
		};
		if( arguments.KeyExists( "type" ) )
			args.type = type;
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
		for( var rowNumber = startRow; rowNumber LTE endRow; rowNumber++ ){
			for( var columnNumber = startColumn; columnNumber LTE endColumn; columnNumber++ )
				setCellValue( workbook, value, rowNumber, columnNumber );
		}
	}

	public void function setColumnWidth( required workbook,required numeric column,required numeric width ){
		var columnIndex = ( column-1 );
		getActiveSheet( workbook ).setColumnWidth( JavaCast( "int", columnIndex ), JavaCast( "int", ( width * 256 ) ) );
	}

	public void function setFooter(
		required workbook
		,string leftFooter=""
		,string centerFooter=""
		,string rightFooter=""
	){
		if( !centerFooter.isEmpty() )
			getActiveSheet( workbook ).getFooter().setCenter( JavaCast( "string", centerFooter ) );
		if( !leftFooter.isEmpty() )
			getActiveSheet( workbook ).getFooter().setleft( JavaCast( "string", leftFooter ) );
		if( !rightFooter.isEmpty() )
			getActiveSheet( workbook ).getFooter().setright( JavaCast( "string", rightFooter ) );
	}

	public void function setHeader(
		required workbook
		,string leftHeader=""
		,string centerHeader=""
		,string rightHeader=""
	){
		if( !centerHeader.isEmpty() )
			getActiveSheet( workbook ).getHeader().setCenter( JavaCast( "string", centerHeader ) );
		if( !leftHeader.isEmpty() )
			getActiveSheet( workbook ).getHeader().setleft( JavaCast( "string", leftHeader ) );
		if( !rightHeader.isEmpty() )
			getActiveSheet( workbook ).getHeader().setright( JavaCast( "string", rightHeader ) );
	}

	public void function setReadOnly( required workbook, required string password ){
		if( isXmlFormat( workbook ) )
			throw( type=exceptionType, message="setReadOnly not supported for XML workbooks", detail="The setReadOnly() method only works on binary 'xls' workbooks." );
		// writeProtectWorkbook takes both a user name and a password, just making up a user name
		workbook.writeProtectWorkbook( JavaCast( "string", password ), JavaCast( "string", "user" ) );
	}

	public void function setRepeatingColumns(
		required workbook
		,required string columnRange
	){
		columnRange = columnRange.Trim();
		if( !IsValid( "regex",columnRange,"[A-Za-z]:[A-Za-z]" ) )
			throw( type=exceptionType, message="Invalid columnRange argument", detail="The 'columnRange' argument should be in the form 'A:B'" );
		var cellRangeAddress = loadPoi( "org.apache.poi.ss.util.CellRangeAddress" ).valueOf( JavaCast( "String", columnRange ) );
		getActiveSheet( workbook ).setRepeatingColumns( cellRangeAddress );
	}

	public void function setRepeatingRows(
		required workbook
		,required string rowRange
	){
		rowRange = rowRange.Trim();
		if( !IsValid( "regex",rowRange,"\d+:\d+" ) )
			throw( type=exceptionType, message="Invalid rowRange argument", detail="The 'rowRange' argument should be in the form 'n:n', e.g. '1:5'" );
		var cellRangeAddress=loadPoi( "org.apache.poi.ss.util.CellRangeAddress" ).valueOf( JavaCast( "String", rowRange ) );
		getActiveSheet( workbook ).setRepeatingRows( cellRangeAddress );
	}

	public void function setRowHeight( required workbook,required numeric row,required numeric height ){
		var rowIndex = ( row -1 );
		getActiveSheet( workbook ).getRow( JavaCast( "int", rowIndex ) ).setHeightInPoints( JavaCast( "int", height ) );
	}

	public void function setSheetPrintOrientation( required workbook, required string mode, string sheetName, numeric sheetNumber ){
		if( !ListFindNoCase( "landscape,portrait", mode ) )
			throw( type=exceptionType, message="Invalid mode argument", detail="#mode# is not a valid 'mode' argument. Use 'portrait' or 'landscape'" );
		var sheetNameSupplied = ( arguments.KeyExists( "sheetName" ) AND Len( sheetName ) );
		if( sheetNameSupplied AND arguments.KeyExists( "sheetNumber" ) )
			throw( type=exceptionType, message="Invalid arguments", detail="Specify either a sheetName or sheetNumber, not both" );
		var setToLandscape = ( LCase( mode ) IS "landscape" );
		if( sheetNameSupplied )
			var sheet = getSheetByName( workbook, sheetName );
		else if( arguments.KeyExists( "sheetNumber" ) )
			var sheet = getSheetByNumber( workbook, sheetNumber );
		else
			var sheet = getActiveSheet( workbook );
		sheet.getPrintSetup().setLandscape( JavaCast( "boolean", setToLandscape ) );
	}

	public void function shiftColumns( required workbook, required numeric start, numeric end=start, numeric offset=1 ){
		if( start LTE 0 )
			throw( type=exceptionType, message="Invalid start value", detail="The start value must be greater than or equal to 1" );
		if( arguments.KeyExists( "end" ) AND ( end LTE 0 OR end LT start ) )
			throw( type=exceptionType, message="Invalid end value", detail="The end value must be greater than or equal to the start value" );
		var rowIterator = getActiveSheet( workbook ).rowIterator();
		var startIndex = start-1;
		var endIndex = arguments.KeyExists( "end" )? end-1: startIndex;
		while( rowIterator.hasNext() ){
			var row = rowIterator.next();
			if( offset GT 0 ){
				for( var i = endIndex; i GTE startIndex; i-- ){
					var tempCell = row.getCell( JavaCast( "int", i ) );
					var cell = createCell( row, i+offset );
					if( !IsNull( tempCell ) ){
						setCellValueAsType( workbook, cell, getCellValueAsType( workbook, tempCell ) );
						cell.setCellStyle( tempCell.getCellStyle() );
					}
				}
			}
			else {
				for( var i = startIndex; i LTE endIndex; i++ ){
					var tempCell = row.getCell( JavaCast( "int", i ) );
					var cell = createCell( row, i+offset );
					if( !IsNull( tempCell ) ){
						setCellValueAsType( workbook, cell, getCellValueAsType( workbook, tempCell ) );
						cell.setCellStyle( tempCell.getCellStyle() );
					}
				}
			}
		}
		// clean up any columns that need to be deleted after the shift
		var numberColsShifted = ( ( endIndex-startIndex ) +1 );
		var numberColsToDelete = Abs( offset );
		if( numberColsToDelete GT numberColsShifted )
			numberColsToDelete = numberColsShifted;
		if( offset GT 0 ){
			var stopValue = ( ( startIndex + numberColsToDelete ) -1 );
			for( var i = startIndex; i LTE stopValue; i++ )
				deleteColumn( workbook, ( i +1 ) );
			return;
		}
		var stopValue = ( ( endIndex - numberColsToDelete ) +1 );
		for( var i = endIndex; i GTE stopValue; i-- )
			deleteColumn( workbook, ( i +1 ) );
	}

	public void function shiftRows( required workbook,required numeric start,numeric end=start,numeric offset=1 ){
		getActiveSheet( workbook ).shiftRows(
			JavaCast( "int", ( arguments.start - 1 ) )
			,JavaCast( "int", ( arguments.end - 1 ) )
			,JavaCast( "int", arguments.offset )
		);
	}

	public void function showColumn( required workbook, required numeric column ){
		toggleColumnHidden( workbook, column,false );
	}

	public void function write( required workbook, required string filepath, boolean overwrite=false, string password, string algorithm="agile" ){
		if( !overwrite AND FileExists( filepath ) )
			throw( type=exceptionType, message="File already exists", detail="The file path specified already exists. Use 'overwrite=true' if you wish to overwrite it." );
		var passwordProtect = ( arguments.KeyExists( "password" ) AND !password.Trim().IsEmpty() );
		if( passwordProtect AND !engineSupportsEncryption )
			throw( type=exceptionType, message="Password protection is not supported for Adobe ColdFusion", detail="Password protection currently only works in Lucee, not ColdFusion" );
		if( passwordProtect AND isBinaryFormat( workbook ) )
			throw( type=exceptionType, message="Whole file password protection is not supported for binary workbooks", detail="Password protection only works with XML ('xlsx') workbooks." );
		lock name="#filepath#" timeout=5{
			var outputStream = CreateObject( "java", "java.io.FileOutputStream" ).init( filepath );
		}
		try{
			lock name="#filepath#" timeout=5{
				workbook.write( outputStream );
			}
			outputStream.flush();
		}
		finally{
			// always close the stream. otherwise file may be left in a locked state if an unexpected error occurs
			if( local.KeyExists( "outputStream" ) )
				outputStream.close();
		}
		if( passwordProtect )
			encryptFile( filepath, password, algorithm );
	}

	/* END PUBLIC API */

	/* PRIVATE METHODS */

	private void function addInfoBinary( required workbook, required struct info ){
		workbook.createInformationProperties(); // creates the following if missing
		var documentSummaryInfo = workbook.getDocumentSummaryInformation();
		var summaryInfo = workbook.getSummaryInformation();
		for( var key in info ){
			var value = JavaCast( "string", info[ key ] );
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
		var documentProperties = workbook.getProperties().getExtendedProperties().getUnderlyingProperties();
		var coreProperties = workbook.getProperties().getCoreProperties();
		for( var key in info ){
			var value=JavaCast( "string", info[ key ] );
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

	private void function addRowToSheetData( required workbook, required struct sheet, required numeric rowIndex, boolean includeRichTextFormatting=false ){
		if( ( rowIndex EQ sheet.headerRowIndex ) AND !sheet.includeHeaderRow )
			return;
		var rowData = [];
		var row = sheet.object.GetRow( JavaCast( "int", rowIndex ) );
		if( IsNull( row ) ){
			if( sheet.includeBlankRows )
				sheet.data.Append( rowData );
			return;
		}
		if( rowIsEmpty( row ) AND !sheet.includeBlankRows )
			return;
		rowData = getRowData( workbook, row, sheet.columnRanges, includeRichTextFormatting );
		sheet.data.Append( rowData );
		if( !sheet.columnRanges.Len() ){
			var rowColumnCount = row.GetLastCellNum();
			sheet.totalColumnCount = Max( sheet.totalColumnCount, rowColumnCount );
		}
	}

	private struct function binaryInfo( required workbook ){
		var documentProperties = workbook.getDocumentSummaryInformation();
		var coreProperties = workbook.getSummaryInformation();
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
		var rowIndex = ( rowNumber -1 );
		var columnIndex = ( columnNumber -1 );
		var checkRow = getActiveSheet( workbook ).getRow( JavaCast( "int", rowIndex ) );
		return !IsNull( checkRow ) AND !IsNull( checkRow.getCell( JavaCast( "int", columnIndex ) ) );
	}

	private numeric function columnCountFromRanges( required array ranges ){
		var result = 0;
		for( var thisRange in ranges ){
			for( var i = thisRange.startAt; i LTE thisRange.endAt; i++ )
				result++;
		}
		return result;
	}

	private any function createCell( required row, numeric cellNum=arguments.row.getLastCellNum(), overwrite=true ){
		/* get existing cell (if any)  */
		var cell = row.getCell( JavaCast( "int", cellNum ) );
		if( overwrite AND !IsNull( cell ) )
			arguments.row.removeCell( cell );/* forcibly remove the existing cell  */
		if( overwrite OR IsNull( cell ) )
			cell = row.createCell( JavaCast( "int", cellNum ) );/* create a brand new cell  */
		return cell;
	}

	private any function createRow( required workbook, numeric rowNum=getNextEmptyRow( workbook ), boolean overwrite=true ){
		/* get existing row (if any)  */
		var row = getActiveSheet( workbook ).getRow( JavaCast( "int", rowNum ) );
		if( overwrite AND !IsNull( row ) )
			getActiveSheet( workbook ).removeRow( row ); /* forcibly remove existing row and all cells  */
		if( overwrite OR IsNull( getActiveSheet( workbook ).getRow( JavaCast( "int", rowNum ) ) ) ){
			try{
				row = getActiveSheet( workbook ).createRow( JavaCast( "int", rowNum ) );
			}
			catch( java.lang.IllegalArgumentException exception ){
				if( exception.message.FindNoCase( "Invalid row number (65536)" ) )
					throw( type=exceptionType, message="Too many rows", detail="Binary spreadsheets are limited to 65535 rows. Consider using an XML format spreadsheet instead." );
				else
					rethrow;
			}
		}
		return row;
	}

	private any function createWorkBook( required string sheetName, boolean useXmlFormat=false ){
		validateSheetName( sheetName );
		var className = useXmlFormat? "org.apache.poi.xssf.usermodel.XSSFWorkbook": "org.apache.poi.hssf.usermodel.HSSFWorkbook";
		return loadPoi( className ).init();
	}

	private any function decryptFile( required string filepath, required string password ){
		var isBinaryFile = ( filepath.ListLast( "." ) IS "xls" );
		if( isBinaryFile )
			throw( type=exceptionType, message="Invalid file type", detail="The library only supports opening encrypted XML (.xlsx) spreadsheets. This file appears to be a binary (.xls) spreadsheet." );
		lock name="#filepath#" timeout=5 {
			try{
				var file = CreateObject( "java", "java.io.File" ).init( filepath );
				var fs = loadPoi( "org.apache.poi.poifs.filesystem.NPOIFSFileSystem" ).init( file );
				if( requiresJavaLoader )
					/* See encryptFile() for explanation of the following line */
					var info = New decryption( server[ poiLoaderName ], fs ).loadInfoWithSwitchedContextLoader();
				else
					var info = loadPoi( "org.apache.poi.poifs.crypt.EncryptionInfo" ).init( fs );;
				var decryptor = loadPoi( "org.apache.poi.poifs.crypt.Decryptor" ).getInstance( info );
				if( decryptor.verifyPassword( password ) )
					return loadPoi( "org.apache.poi.xssf.usermodel.XSSFWorkbook" ).init( decryptor.getDataStream( fs ) );
				throw( type=exceptionType,message="Invalid password",detail="The file cannot be read because the password is incorrect." );
			}
			catch( org.apache.poi.poifs.filesystem.NotOLE2FileException exception ){
				throw( type=exceptionType, message="Invalid spreadsheet file", detail="The file #filepath# does not appear to be a spreadsheet" );
			}
			finally{
				if( local.KeyExists( "fs" ) )
					fs.close();
			}
		}
	}

	private query function deleteHiddenColumnsFromQuery( required sheet, required query result ){
		var startIndex = ( sheet.totalColumnCount -1 );
		for( var colIndex = startIndex; colIndex GTE 0; colIndex-- ){
			if( !sheet.object.isColumnHidden( JavaCast( "int", colIndex ) ) )
				continue;
			var columnNumber = ( colIndex +1 );
			result = _QueryDeleteColumn( result, sheet.columnNames[ columnNumber ] );
			sheet.totalColumnCount--;
			sheet.columnNames.deleteAt( columnNumber );
		}
		return result;
	}

	private void function deleteSheetAtIndex( required workbook, required numeric sheetIndex ){
		workbook.removeSheetAt( JavaCast( "int", sheetIndex ) );
	}

	private string function detectValueDataType( required value ){
		// Numeric must precede date test
		// Golden default rule: treat numbers with leading zeros as STRINGS: not numbers (lucee) or dates (ACF);
		if( REFind( "^0[\d]+", value ) )
			return "string";
		if( IsNumeric( value ) )
			return "numeric";
		if( _isDate( value ) )
			return "date";
		if( !Len( Trim( value ) ) )
			return "blank";
		return "string";
	}

	private void function downloadBinaryVariable( required binaryVariable, required string filename, required contentType ){
		cfheader( name="Content-Disposition", value='attachment; filename="#filename#"' );
		cfcontent( type=contentType, variable="#binaryVariable#", reset="true" );
	}

	private void function encryptFile( required string filepath, required string password, required string algorithm ){
		/* See http://poi.apache.org/encryption.html */
		/* NB: Not all spreadsheet programs support this type of encryption */
		lock name="#filepath#" timeout=5 {
			try{
				var fs = loadPoi( "org.apache.poi.poifs.filesystem.POIFSFileSystem" );
				if( requiresJavaLoader )
					/*
						Need to ensure our poiLoader is maintained as the "contextLoader" so that when POI objects load other POI objects, they find them. Otherwise Lucee's loader would be used, which isn't aware of our POI library. JavaLoader supports this via a complicated "mixin" procedure: https://github.com/markmandel/JavaLoader/wiki/Switching-the-ThreadContextClassLoader
					*/
					var info = New encryption( server[ poiLoaderName ], algorithm ).loadInfoWithSwitchedContextLoader();
				else {
					var mode = loadPoi( "org.apache.poi.poifs.crypt.EncryptionMode" );
					switch( algorithm ){
						case "agile":
							var info = loadPoi( "org.apache.poi.poifs.crypt.EncryptionInfo" ).init( mode.agile );
							break;
						case "standard":
							var info = loadPoi( "org.apache.poi.poifs.crypt.EncryptionInfo" ).init( mode.standard );
							break;
						case "binaryRC4":
							var info = loadPoi( "org.apache.poi.poifs.crypt.EncryptionInfo" ).init( mode.binaryRC4 );
							break;
					}
				}
				var encryptor = info.getEncryptor();
				encryptor.confirmPassword( JavaCast( "string", password ) );
				var opcAccess = loadPoi( "org.apache.poi.openxml4j.opc.PackageAccess" );
				try{
					var file = CreateObject( "java", "java.io.File" ).init( filepath );
					var opc = loadPoi( "org.apache.poi.openxml4j.opc.OPCPackage" ).open( file, opcAccess.READ_WRITE );
					var encryptedStream = encryptor.getDataStream( fs );
					opc.save( encryptedStream );
				}
				finally{
					if( local.KeyExists( "opc" ) )
						opc.close();
				}
				try{
					var outputStream = CreateObject( "java", "java.io.FileOutputStream" ).init( filepath );
					fs.writeFilesystem( outputStream );
					outputStream.flush();
				}
				finally{
					// always close the stream. otherwise file may be left in a locked state if an unexpected error occurs
					if( local.KeyExists( "outputStream" ) )
						outputStream.close();
				}
			}
			finally{
				if( local.KeyExists( "fs" ) )
					fs.close();
			}
		}
	}

	private numeric function estimateColumnWidth( required workbook, required any value ){
		/* Estimates approximate column width based on cell value and default character width. */
		/*
		"Excel bases its measurement of column widths on the number of digits (specifically, the number of zeros) in the column, using the Normal style font."
		This function approximates the column width using the number of characters and the default character width in the normal font. POI expresses the width in 1/256 of Excel's character unit. The maximum size in POI is: (255 * 256)
		*/
		var defaultWidth = getDefaultCharWidth( workbook );
		var numOfChars = Len( arguments.value );
		var width = ( numOfChars * defaultWidth +5 ) / ( defaultWidth * 256 );
	    // Do not allow the size to exceed POI's maximum
		return Min( width, ( 255 * 256 ) );
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
			if( !REFind( rangeTest, thisRange ) )
				throw( type=exceptionType, message="Invalid range value", detail="The range value '#thisRange#' is not valid." );
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
		var result = input.REReplace( "[#charsToRemove#]+", "", "ALL" ).Left( 255 );
		if( result.isEmpty() )
			return "renamed"; // in case all chars have been replaced (unlikely but possible)
		return result;
	}

	private void function doFillMergedCellsWithVisibleValue( required workbook, required sheet ){
		if( !sheetHasMergedRegions( sheet ) )
			return;
		for( var regionIndex = 0; regionIndex LT sheet.getNumMergedRegions(); regionIndex++ ){
			var region = sheet.getMergedRegion( regionIndex );
			var regionStartRowNumber = ( region.getFirstRow() +1 );
			var regionEndRowNumber = ( region.getLastRow() +1 );
			var regionStartColumnNumber = ( region.getFirstColumn() +1 );
			var regionEndColumnNumber = ( region.getLastColumn() +1 );
			var visibleValue = getCellValue( workbook, regionStartRowNumber, regionStartColumnNumber );
			setCellRangeValue( workbook, visibleValue, regionStartRowNumber, regionEndRowNumber, regionStartColumnNumber, regionEndColumnNumber );
		}
	}

	private string function generateUniqueSheetName( required workbook ){
		/* Generates a unique sheet name (Sheet1, Sheet2, etecetera). */
		var startNumber = ( workbook.getNumberOfSheets() +1 );
		var maxRetry = ( startNumber +250 );
		for( var sheetNumber = startNumber; sheetNumber LTE maxRetry; sheetNumber++ ){
			var proposedName = "Sheet" & sheetNumber;
			if( !sheetExists( workbook,proposedName ) )
				return proposedName;
		}
		/* this should never happen. but if for some reason it did, warn the action failed and abort */
		throw( type=exceptionType, message="Unable to generate name", detail="Unable to generate a unique sheet name" );
	}

	private any function getActiveSheet( required workbook ){
		return workbook.getSheetAt( JavaCast( "int", workbook.getActiveSheetIndex() ) );
	}

	private any function getActiveSheetName( required workbook ){
		return getActiveSheet( workbook ).getSheetName();
	}

	private numeric function getAWTFontStyle( required any poiFont ){
		var font = loadPOI( "java.awt.Font" );
		var isBold = poiFont.getBold();
		if( isBold && arguments.poiFont.getItalic() )
	  	return BitOr( font.BOLD, font.ITALIC );
		if( isBold )
			return font.BOLD;
		if( poiFont.getItalic() )
			return font.ITALIC;
		return font.PLAIN;
	}

	private any function getCellAt( required workbook, required numeric rowNumber, required numeric columnNumber ){
		if( !cellExists( argumentCollection=arguments ) )
			throw( type=exceptionType, message="Invalid cell", detail="The requested cell [#rowNumber#,#columnNumber#] does not exist in the active sheet" );
		var rowIndex = ( rowNumber -1 );
		var columnIndex = ( columnNumber -1 );
		return getActiveSheet( workbook ).getRow( JavaCast( "int", rowIndex ) ).getCell( JavaCast( "int", columnIndex ) );
	}

	private any function getCellUtil(){
		if( IsNull( variables.cellUtil ) )
			variables.cellUtil = loadPoi( "org.apache.poi.ss.util.CellUtil" );
		return variables.cellUtil;
	}

	private any function getCellValueAsType( required workbook, required cell ){
		/* When getting the value of a cell, it is important to know what type of cell value we are dealing with. If you try to grab the wrong value type, an error might be thrown. For that reason, we must check to see what type of cell we are working with. These are the cell types and they are constants of the cell object itself:

		20170116: In POI 4.0 getCellType() will no longer return an integer, but a CellType enum instead. Shouldn't affect things as we are only using the constants, not the integer literals.

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
			var dateUtil = getDateUtil();
			if( dateUtil.isCellDateFormatted( cell ) ){
				var cellValue = cell.getDateCellValue();
				if( DateCompare( "1899-12-31", cellValue, "d" ) EQ 0 ) // TIME
					return getFormatter().formatCellValue( cell );//return as a time formatted string to avoid default epoch date 1899-12-31
				return cellValue;
			}
			return cell.getNumericCellValue();
		}
		if( cellType EQ cell.CELL_TYPE_FORMULA ){
			var formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
			try{
				return getFormatter().formatCellValue( cell, formulaEvaluator );
			}
			catch( any exception ){
				throw( type=exceptionType, message="Failed to run formula", detail="There is a problem with the formula in sheet #cell.getSheet().getSheetName()# row #( cell.getRowIndex() +1 )# column #( cell.getColumnIndex() +1 )#");
			}
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

	private any function getDateUtil(){
		if( IsNull( variables.dateUtil ) )
			variables.dateUtil = loadPoi( "org.apache.poi.ss.usermodel.DateUtil" );
		return variables.dateUtil;
	}

	private string function getDateTimeValueFormat( required any value ){
		/* Returns the default date mask for the given value: DATE (only), TIME (only) or TIMESTAMP */
		var dateTime = ParseDateTime( value );
		var dateOnly = CreateDate( Year( dateTime ), Month( dateTime ), Day( dateTime ) );
		if( DateCompare( value, dateOnly, "s" ) EQ 0 )
			return variables.dateFormats.DATE;
		if( DateCompare( "1899-12-30", dateOnly, "d" ) EQ 0 )
			return variables.dateFormats.TIME;
		return variables.dateFormats.TIMESTAMP;
	}

	private numeric function getDefaultCharWidth( required workbook ){
		/* Estimates the default character width using Excel's 'Normal' font */
		/* this is a compromise between hard coding a default value and the more complex method of using an AttributedString and TextLayout */
		var defaultFont = workbook.getFontAt( 0 );
		var style = getAWTFontStyle( defaultFont );
		var font = loadPOI( "java.awt.Font" );
		var javaFont = font.init( defaultFont.getFontName(), style, defaultFont.getFontHeightInPoints() );
		// this works
		var transform = CreateObject( "java", "java.awt.geom.AffineTransform" );
		var fontContext = CreateObject( "java", "java.awt.font.FontRenderContext" ).init( transform, true, true );
		var bounds = javaFont.getStringBounds( "0", fontContext );
		return bounds.getWidth();
	}

	private numeric function getFirstRowNum( required workbook ){
		var firstRow = getActiveSheet( workbook ).getFirstRowNum();
		if( firstRow EQ 0 AND getActiveSheet( workbook ).getPhysicalNumberOfRows() EQ 0 )
			return -1;
		return firstRow;
	}

	private any function getFormatter(){
		/* Returns cell formatting utility object ie org.apache.poi.ss.usermodel.DataFormatter */
		if( IsNull( variables.dataFormatter ) )
			variables.dataFormatter = loadPOI( "org.apache.poi.ss.usermodel.DataFormatter" ).init();
		return dataFormatter;
	}

	private struct function getJavaColorRGB( required string colorName ){
		/* Returns a struct containing RGB values from java.awt.Color for the color name passed in */
		var findColor = colorName.Trim().UCase();
		var color = CreateObject( "Java", "java.awt.Color" );
		if( IsNull( color[ findColor ] ) OR !IsInstanceOf( color[ findColor ], "java.awt.Color" ) )//don't use member functions on color
			throw( type=exceptionType, message="Invalid color", detail="The color provided (#colorName#) is not valid." );
		color = color[ findColor ];
		var colorRGB = {
			red: color.getRed()
			,green: color.getGreen()
			,blue: color.getBlue()
		};
		return colorRGB;
	}

	private numeric function getLastRowNum( required workbook ){
		var lastRow = getActiveSheet( workbook ).getLastRowNum();
		if( lastRow EQ 0 AND getActiveSheet( workbook ).getPhysicalNumberOfRows() EQ 0 )
			return -1; //The sheet is empty. Return -1 instead of 0
		return lastRow;
	}

	private numeric function getNextEmptyRow( workbook ){
		return ( getLastRowNum( workbook ) +1 );
	}

	private array function getQueryColumnFormats( required workbook, required query query ){
		/* extract the query columns and data types  */
		var formatter	= workbook.getCreationHelper().createDataFormat();
		var metadata = GetMetaData( query );
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

	private array function getRowData( required workbook, required row, array columnRanges=[], boolean includeRichTextFormatting=false ){
		var result = [];
		if( !columnRanges.Len() ){
			var columnRange = {
				startAt: 1
				,endAt: row.GetLastCellNum()
			};
			arguments.columnRanges = [ columnRange ];
		}
		for( var thisRange in columnRanges ){
			for( var i = thisRange.startAt; i LTE thisRange.endAt; i++ ){
				var colIndex = ( i-1 );
				var cell = row.GetCell( JavaCast( "int", colIndex ) );
				if( IsNull( cell ) ){
					result.Append( "" );
					continue;
				}
				var cellValue = getCellValueAsType( workbook, cell );
				if( includeRichTextFormatting AND ( cell.GetCellType() EQ cell.CELL_TYPE_STRING ) )
					cellValue = richStringCellValueToHtml( workbook, cell,cellValue );
				result.Append( cellValue );
			}
		}
		return result;
	}

	private any function getSheetByName( required workbook, required string sheetName ){
		validateSheetExistsWithName( workbook, sheetName );
		return workbook.getSheet( JavaCast( "string", sheetName ) );
	}

	private any function getSheetByNumber( required workbook, required numeric sheetNumber ){
		validateSheetNumber( workbook, sheetNumber );
		var sheetIndex = ( sheetNumber -1 );
		return workbook.getSheetAt( sheetIndex );
	}

	private numeric function getSheetIndexFromName( required workbook, required string sheetName ){
		//returns -1 if non-existent
		return workbook.getSheetIndex( JavaCast( "string", sheetName ) );
	}

	private any function initializeCell( required workbook, required numeric rowNumber, required numeric columnNumber ){
		var rowIndex = JavaCast( "int", ( rowNumber -1 ) );
		var columnIndex = JavaCast( "int", ( columnNumber -1 ) );
		var rowObject = getCellUtil().getRow( rowIndex, getActiveSheet( workbook ) );
		var cellObject = getCellUtil().getCell( rowObject, columnIndex );
		return cellObject;
	}

	private boolean function isCsvOrTextFile( required string path ){
		var contentType = FileGetMimeType( path ).ListLast( "/" );
		return ListFindNoCase( "plain,csv", contentType );//Lucee=text/plain ACF=text/csv
	}

	private boolean function isDateObject( required input ){
		return input.getClass().getName() IS "java.util.Date";
	}

	private boolean function isString( required input ){
		return input.getClass().getName() IS "java.lang.String";
	}

	private array function getPoiJarPaths(){
		var libPath = GetDirectoryFromPath( GetCurrentTemplatePath() ) & "lib/";
		return DirectoryList( libPath );
	}

	private function loadPoi( required string javaclass ){
		if( !requiresJavaLoader ){
			if( engineSupportsDynamicClassLoading ){
				poiClassesLastLoadedVia = "Dynamic loading from the lib folder";
				return CreateObject( "java", javaclass, getPoiJarPaths() );
			}
			// Else *the correct* POI jars must be in the class path and any older versions *removed*
			try{
				poiClassesLastLoadedVia = "The java class path";
				return CreateObject( "java", javaclass );
			}
			catch( any exception ){
				poiClassesLastLoadedVia = "JavaLoader";
				return loadPoiUsingJavaLoader( javaclass );
			}
		}
		poiClassesLastLoadedVia = "JavaLoader";
		return loadPoiUsingJavaLoader( javaclass );
	}

	private void function handleInvalidSpreadsheetFile( required string path ){
		var detail = "The file #path# does not appear to be a binary or xml spreadsheet.";
		if( isCsvOrTextFile( path ) )
			detail &= " It may be a CSV file, in which case use 'csvToQuery()' to read it";
		throw( type="cfsimplicity.lucee.spreadsheet.invalidFile", message="Invalid spreadsheet file", detail=detail );
	}

	private function loadPoiUsingJavaLoader( required string javaclass ){
		if( !server.KeyExists( poiLoaderName ) )
			server[ poiLoaderName ] = CreateObject( "component", javaLoaderDotPath ).init( loadPaths=getPoiJarPaths(), loadColdFusionClassPath=true, trustedSource=true );
		return server[ poiLoaderName ].create( javaclass );
	}

	private void function moveSheet( required workbook, required string sheetName, required string moveToIndex ){
		workbook.setSheetOrder( JavaCast( "String", sheetName ), JavaCast( "int", moveToIndex ) );
	}

	private array function parseRowData( required string line, required string delimiter, boolean handleEmbeddedCommas=true ){
		var elements = ListToArray( arguments.line, arguments.delimiter );
		var potentialQuotes = 0;
		arguments.line = ToString( arguments.line );
		if( arguments.delimiter EQ "," AND arguments.handleEmbeddedCommas )
			potentialQuotes = arguments.line.replaceAll( "[^']", "" ).length();
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
				  finalValue = finalValue.substring( ( startAt +1 ),endAt );
			  values.add( finalValue );
			  buffer.setLength( 0 );
			  isEmbeddedValue = false;
		  }
	  }
	  return values;
	}

	private string function queryToCsv( required query query, numeric headerRow, boolean includeHeaderRow ){
		var result = CreateObject( "Java", "java.lang.StringBuilder" ).init();
		var crlf = Chr( 13 ) & Chr( 10 );
		var columns = _QueryColumnArray( query );
		var hasHeaderRow = ( arguments.KeyExists( "headerRow" ) AND Val( headerRow ) );
		if( hasHeaderRow )
			result.Append( generateCsvRow( columns ) );
		for( var row in query ){
			var rowValues = [];
			for( column in columns )
				rowValues.Append( row[ column ] );
			result.Append( crlf & generateCsvRow( rowValues ) );
		}
		return result.toString().Trim();
	}

	private string function generateCsvRow( required array values, delimiter="," ){
		var result = CreateObject( "Java", "java.lang.StringBuilder" ).init();
		for( var value in values ){
			if( isDateObject( value ) )
				value = DateTimeFormat( value, dateFormats.DATETIME );
			value = Replace( value, '"', '""', "ALL" );//can't use member function in case its a non-string
			result.Append( '#delimiter#"#value#"' );
		}
		return result.toString().substring( 1 );
	}

	private string function queryToHtml( required query query, numeric headerRow, boolean includeHeaderRow ){
		var result = CreateObject( "Java", "java.lang.StringBuilder" ).init();
		var columns = _QueryColumnArray( query );
		var hasHeaderRow = ( arguments.KeyExists( "headerRow" ) AND Val( headerRow ) );
		if( hasHeaderRow ){
			result.Append( "<thead>" );
			result.Append( generateHtmlRow( columns, true ) );
			result.Append( "</thead>" );
		}
		result.Append( "<tbody>" );
		for( var row in query ){
			var rowValues=[];
			for( column in columns )
				rowValues.Append( row[ column ] );
			result.Append( generateHtmlRow( rowValues ) );
		}
		result.Append( "</tbody>" );
		return result.toString();
	}

	private string function generateHtmlRow( required array values, boolean isHeader=false ){
		var result = CreateObject( "Java", "java.lang.StringBuilder" ).init();
		result.Append( "<tr>" );
		var columnTag = isHeader? "th": "td";
		for( var value in values ){
			if( isDateObject( value ) )
				value = DateTimeFormat( value, dateFormats.DATETIME );
			result.Append( "<#columnTag#>#value#</#columnTag#>" );
		}
		result.Append( "</tr>" );
		return result.toString();
	}

	private boolean function rowIsEmpty( required row ){
		for( var i = row.getFirstCellNum(); i LT row.getLastCellNum(); i++ ){
	    var cell = row.getCell( i );
	    if( !IsNull( cell ) && ( cell.getCellType() != cell.CELL_TYPE_BLANK ) )
	      return false;
	  }
	  return true;
	}

	private void function setCellValueAsType( required workbook, required cell, required value, string type ){
		if( !arguments.KeyExists( "type" ) ) //autodetect type
			arguments.type = detectValueDataType( value );
		else if( !ListFindNoCase( "string,numeric,date,boolean,blank", type ) )
			throw( type=exceptionType, message="Invalid data type: '#type#'", detail="The data type must be one of 'string', 'numeric', 'date' 'boolean' or 'blank'." );
		/* Note: To properly apply date/number formatting:
			- cell type must be CELL_TYPE_NUMERIC
			- cell value must be applied as a java.util.Date or java.lang.Double (NOT as a string)
			- cell style must have a dataFormat (datetime values only)
 		*/
		switch( type ){
			case "numeric":
				cell.setCellType( cell.CELL_TYPE_NUMERIC );
				cell.setCellValue( JavaCast( "double", Val( value ) ) );
				return;
			case "date":
				var cellFormat = getDateTimeValueFormat( value );
				var formatter = workbook.getCreationHelper().createDataFormat();
				//Use setCellStyleProperty() which will try to re-use an existing style rather than create a new one for every cell which may breach the 4009 styles per wookbook limit
				getCellUtil().setCellStyleProperty( cell, getCellUtil().DATA_FORMAT, formatter.getFormat( JavaCast( "string", cellFormat ) ) );
				cell.setCellType( cell.CELL_TYPE_NUMERIC );
				/*  Excel's uses a different epoch than CF (1900-01-01 versus 1899-12-30). "Time" only values will not display properly without special handling - */
				if( cellFormat EQ variables.dateFormats.TIME ){
					var dateUtil = getDateUtil();
					value = TimeFormat( value, "HH:MM:SS" );
				 	cell.setCellValue( dateUtil.convertTime( value ) );
				}
				else
					cell.setCellValue( ParseDateTime( value ) );
				return;
			case "boolean":
				cell.setCellType( cell.CELL_TYPE_BOOLEAN );
				cell.setCellValue( JavaCast( "boolean", value ) );
				return;
			case "blank":
				cell.setCellType( cell.CELL_TYPE_BLANK ); //no need to set the value: it will be blank
				return;
		}
		// string
		cell.setCellType( cell.CELL_TYPE_STRING );
		cell.setCellValue( JavaCast( "string", value ) );
	}

	private boolean function sheetExists( required workbook, string sheetName, numeric sheetNumber ){
		validateSheetNameOrNumberWasProvided( argumentCollection=arguments );
		if( arguments.KeyExists( "sheetName" ) )
			arguments.sheetNumber = ( getSheetIndexFromName( workbook, sheetName ) +1 );
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
		var sheet = {
			includeHeaderRow: includeHeaderRow
			,hasHeaderRow: ( arguments.KeyExists( "headerRow" ) AND Val( headerRow ) )
			,includeBlankRows: includeBlankRows
			,columnNames: []
			,columnRanges: []
			,totalColumnCount: 0
		};
		sheet.headerRowIndex = sheet.hasHeaderRow? ( headerRow -1 ): -1;
		if( arguments.KeyExists( "columns" ) ){
			sheet.columnRanges = extractRanges( arguments.columns );
			sheet.totalColumnCount = columnCountFromRanges( sheet.columnRanges );
		}
		if( arguments.KeyExists( "sheetName" ) ){
			validateSheetExistsWithName( workbook,sheetName );
			arguments.sheetNumber = ( getSheetIndexFromName( workbook,sheetName ) +1 );
		}
		sheet.object = getSheetByNumber( workbook, sheetNumber );
		if( fillMergedCellsWithVisibleValue )
			doFillMergedCellsWithVisibleValue( workbook,sheet.object );
		sheet.data=[];
		if( arguments.KeyExists( "rows" ) ){
			var allRanges=extractRanges( arguments.rows );
			for( var thisRange in allRanges ){
				for( var rowNumber = thisRange.startAt; rowNumber LTE thisRange.endAt; rowNumber++ ){
					var rowIndex = ( rowNumber -1 );
					addRowToSheetData( workbook, sheet, rowIndex, includeRichTextFormatting );
				}
			}
		}
		else {
			var lastRowIndex=sheet.object.GetLastRowNum();// zero based
			for( var rowIndex = 0; rowIndex LTE lastRowIndex; rowIndex++ )
				addRowToSheetData( workbook, sheet, rowIndex, includeRichTextFormatting );
		}
		//generate the query columns
		if( arguments.KeyExists( "columnNames" ) AND arguments.columnNames.Len() ){
			arguments.columnNames = arguments.columnNames.ListToArray();
			var specifiedColumnCount = columnNames.Len();
			for( var i = 1; i LTE sheet.totalColumnCount; i++ ){
				// ACF11 elvis operator doesn't work here for some reason. Forced to use longer ternery syntax. IsNull/IsDefined doesn't work either.
				var columnName = ( i LTE specifiedColumnCount )? columnNames[ i ]: "column" & i;
				sheet.columnNames.Append( columnName );
			}
		}
		else if( sheet.hasHeaderRow ){
			var headerRowObject = sheet.object.GetRow( JavaCast( "int", sheet.headerRowIndex ) );
			var rowData = getRowData( workbook, headerRowObject, sheet.columnRanges );
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
		var result = _QueryNew( sheet.columnNames.ToList(), "", sheet.data );
		if( !includeHiddenColumns ){
			result = deleteHiddenColumnsFromQuery( sheet, result );
			if( sheet.totalColumnCount EQ 0 )
				return QueryNew( "" );// all columns were hidden: return a blank query.
		}
		return result;
	}

	private void function toggleColumnHidden( required workbook, required numeric columnNumber, required boolean state ){
		var sheet = getActiveSheet( workbook );
		sheet.setColumnHidden( JavaCast( "int", columnNumber-1 ), JavaCast( "boolean", state ) );
	}

	private void function validateSheetExistsWithName( required workbook,required string sheetName ){
		if( !sheetExists( workbook=workbook, sheetName=sheetName ) )
			throw( type=exceptionType, message="Invalid sheet name [#sheetName#]", detail="The specified sheet was not found in the current workbook." );
	}

	private void function validateSheetNumber( required workbook,required numeric sheetNumber ){
		if( !sheetExists( workbook=workbook, sheetNumber=sheetNumber ) ){
			var sheetCount = workbook.getNumberOfSheets();
			throw( type=exceptionType, message="Invalid sheet number [#sheetNumber#]", detail="The sheetNumber must a whole number between 1 and the total number of sheets in the workbook [#sheetCount#]" );
		}
	}

	private void function validateSheetName( required string sheetName ){
		var poiTool = loadPoi( "org.apache.poi.ss.util.WorkbookUtil" );
		try{
			poiTool.validateSheetName( JavaCast( "String",sheetName ) );
		}
		catch( "java.lang.IllegalArgumentException" exception ){
			throw( type=exceptionType, message="Invalid characters in sheet name", detail=exception.message );
		}
	}

	private void function validateSheetNameOrNumberWasProvided(){
		if( !arguments.KeyExists( "sheetName" ) AND !arguments.KeyExists( "sheetNumber" ) )
			throw( type=exceptionType, message="Missing Required Argument", detail="Either sheetName or sheetNumber must be provided" );
		if( arguments.KeyExists( "sheetName" ) AND arguments.KeyExists( "sheetNumber" ) )
			throw( type=exceptionType, message="Too Many Arguments", detail="Only one argument is allowed. Specify either a sheetName or sheetNumber, not both" );
	}

	private any function workbookFromFile( required string path ){
		// works with both xls and xlsx
		try{
			lock name="#path#" timeout=5 {
				var file = CreateObject( "java", "java.io.FileInputStream" ).init( path );
				var workbook = loadPoi( "org.apache.poi.ss.usermodel.WorkbookFactory" ).create( file );
			}
			return workbook;
		}
		catch( org.apache.poi.openxml4j.exceptions.InvalidFormatException exception ){
			handleInvalidSpreadsheetFile( path );
		}
		catch( org.apache.poi.hssf.OldExcelFormatException exception ){
			throw( type="cfsimplicity.lucee.spreadsheet.OldExcelFormatException", message="Invalid spreadsheet format", detail="The file #path# was saved in a format that is too old. Please save it as an 'Excel 97/2000/XP' file or later." );
		}
		catch( any exception ){
			//For ACF which doesn't return the correct exception types
			if( exception.message CONTAINS "Your InputStream was neither" )
				handleInvalidSpreadsheetFile( path );
			if( exception.message CONTAINS "spreadsheet seems to be Excel 5" )
				throw( type="cfsimplicity.lucee.spreadsheet.OldExcelFormatException", message="Invalid spreadsheet format", detail="The file #path# was saved in a format that is too old. Please save it as an 'Excel 97/2000/XP' file or later." );
			rethrow;
		}
		finally{
			if( local.KeyExists( "file" ) )
				file.close();
		}
	}

	private struct function xmlInfo( required workbook ){
		var documentProperties = workbook.getProperties().getExtendedProperties().getUnderlyingProperties();
		var coreProperties = workbook.getProperties().getCoreProperties();
		return {
			author: coreProperties.getCreator()?:""
			,category: coreProperties.getCategory()?:""
			,comments: coreProperties.getDescription()?:""
			,creationDate: coreProperties.getCreated()?:""
			,lastEdited: coreProperties.getModified()?:""
			,subject: coreProperties.getSubject()?:""
			,title: coreProperties.getTitle()?:""
			,lastAuthor: coreProperties.getUnderlyingProperties().getLastModifiedByProperty().getValue()?:""
			,keywords: coreProperties.getKeywords()?:""
			,lastSaved: ""// not available in xml
			,manager: documentProperties.getManager()?:""
			,company: documentProperties.getCompany()?:""
		};
	}

	/* Formatting */

	private string function richStringCellValueToHtml( required workbook, required cell, required cellValue ){
		var richTextValue = cell.getRichStringCellValue();
		var totalRuns = richTextValue.numFormattingRuns();
		var baseFont = cell.getCellStyle().getFont( workbook );
		if( totalRuns EQ 0  )
			return baseFontToHtml( workbook, cellValue, baseFont );
		// Runs never start at the beginning: the string before the first run is always in the baseFont format
		var startOfFirstRun = richTextValue.getIndexOfFormattingRun( 0 );
		var initialContents = cellValue.Mid( 1, startOfFirstRun );//before the first run
		var initialHtml = baseFontToHtml( workbook, initialContents, baseFont );
		var result = CreateObject( "Java", "java.lang.StringBuilder" ).init();
		result.Append( initialHtml );
		var endOfCellValuePosition = cellValue.Len();
		for( var runIndex = 0; runIndex LT totalRuns; runIndex++ ){
			var run = {};
			run.index = runIndex;
			run.number = ( runIndex +1 );
			run.font = workbook.getFontAt( richTextValue.getFontOfFormattingRun( runIndex ) );
			run.css = runFontToHtml( workbook, baseFont, run.font );
			run.isLast = ( run.number EQ totalRuns );
			run.startPosition = ( richTextValue.getIndexOfFormattingRun( runIndex ) +1 );
			run.endPosition = run.isLast? endOfCellValuePosition: richTextValue.getIndexOfFormattingRun( ( runIndex +1 ) );
			run.length = ( ( run.endPosition +1 ) -run.startPosition );
			run.content = cellValue.Mid( run.startPosition, run.length );
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
		if( Compare( runFont.getBold(), baseFont.getBold() ) )
			cssStyles.Append( fontStyleToCss( "bold", runFont.getBold() ) );
		/* color */
		if( Compare( runFont.getColor(), baseFont.getColor() ) AND !fontColorIsBlack( runFont.getColor() ) )
			cssStyles.Append( fontStyleToCss( "color", runFont.getColor(), workbook ) );
		/* italic */
		if( Compare( runFont.getItalic(), baseFont.getItalic() ) )
			cssStyles.Append( fontStyleToCss( "italic", runFont.getItalic() ) );
		/* underline/strike */
		if( Compare( runFont.getStrikeout(), baseFont.getStrikeout() ) OR Compare( runFont.getUnderline(), baseFont.getUnderline() ) ){
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
		if( baseFont.getBold() )
			cssStyles.Append( fontStyleToCss( "bold", true ) );
		/* color */
		if( !fontColorIsBlack( baseFont.getColor() ) )
			cssStyles.Append( fontStyleToCss( "color", baseFont.getColor(), workbook ) );
		/* italic */
		if( baseFont.getItalic() )
			cssStyles.Append( fontStyleToCss( "italic", true ) );
		/* underline/strike */
		if( baseFont.getStrikeout() OR baseFont.getUnderline() ){
			var decorationValue	=	[];
			if( baseFont.getStrikeout() )
				decorationValue.Append( "line-through" );
			if( baseFont.getUnderline() )
				decorationValue.Append( "underline" );
			cssStyles.Append( fontStyleToCss( "decoration", decorationValue.ToList( " " ) ) );
		}
		cssStyles = cssStyles.toString();
		if( cssStyles.IsEmpty() )
			return contents;
		return "<span style=""#cssStyles#"">#contents#</span>";
	}

	private string function fontStyleToCss( required string styleType, required any styleValue, workbook ){
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
					throw( type=exceptionType, message="The 'workbook' argument is required when generating color css styles" );
				//http://ragnarock99.blogspot.co.uk/2012/04/getting-hex-color-from-excel-cell.html
				var rgb = workbook.getCustomPalette().getColor( styleValue ).getTriplet();
				var javaColor = CreateObject( "Java", "java.awt.Color" ).init( JavaCast( "int", rgb[ 1 ] ), JavaCast( "int", rgb[ 2 ] ), JavaCast( "int", rgb[ 3 ] ) );
				var hex	=	CreateObject( "Java", "java.lang.Integer" ).toHexString( javaColor.getRGB() );
				hex = hex.subString( 2, hex.length() );
				return "color:##" & hex & ";";
			case "italic":
				return "font-style:" & ( styleValue? "italic;": "normal;" );
			case "decoration":
				return "text-decoration:#styleValue#;";//need to pass desired combination of "underline" and "line-through"
		}
		throw( type=exceptionType, message="Unrecognised style for css conversion" );
	}

	private boolean function fontColorIsBlack( required fontColor ){
		return ( fontColor IS 8 ) OR ( fontColor IS 32767 );
	}

	private any function buildCellStyle( required workbook, required struct format ){
		/*  TODO: Reuse styles  */
		var cellStyle = workbook.createCellStyle();
		var formatter = workbook.getCreationHelper().createDataFormat();
		var font = 0;
		var formatIndex = 0;
		/*
			Valid keys of the format struct are:
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
			var settingValue = format[ setting ];
			switch( setting ){
				case "alignment":
					var alignment = cellStyle.getAlignmentEnum()[ JavaCast( "string", UCase( settingValue ) ) ];
					cellStyle.setAlignment( alignment );
				break;
				case "bold":
					font = cloneFont( workbook,workbook.getFontAt( cellStyle.getFontIndex() ) );
					font.setBold( JavaCast( "boolean", settingValue ) );
					cellStyle.setFont( font );
				break;
				case "bottomborder":
					var borderStyle = cellStyle.getBorderBottomEnum()[ JavaCast( "string", UCase( settingValue ) ) ];
					cellStyle.setBorderBottom( borderStyle );
				break;
				case "bottombordercolor":
					cellStyle.setBottomBorderColor( getColor( workbook, settingValue ) );
				break;
				case "color":
					font = cloneFont( workbook, workbook.getFontAt( cellStyle.getFontIndex() ) );
					font.setColor( getColor( workbook, settingValue ) );
					cellStyle.setFont( font );
				break;
				case "dataformat":
					cellStyle.setDataFormat( formatter.getFormat( JavaCast( "string", settingValue ) ) );
				break;
				case "fgcolor":
					cellStyle.setFillForegroundColor( getColor( workbook, settingValue ) );
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
					font = cloneFont( workbook, workbook.getFontAt( cellStyle.getFontIndex() ) );
					font.setFontName( JavaCast( "string", settingValue ) );
					cellStyle.setFont( font );
				break;
				case "fontsize":
					font = cloneFont( workbook,workbook.getFontAt( cellStyle.getFontIndex() ) );
					font.setFontHeightInPoints( JavaCast( "int", settingValue ) );
					cellStyle.setFont( font );
				break;
				/*  TODO: Doesn't seem to do anything */
				case "hidden":
					cellStyle.setHidden( JavaCast( "boolean", settingValue ) );
				break;
				/*  TODO: Doesn't seem to do anything */
				case "indent":
					cellStyle.setIndention( JavaCast( "int", settingValue ) );
				break;
				case "italic":
					font = cloneFont( workbook,workbook.getFontAt( cellStyle.getFontIndex ( ) ) );
					font.setItalic( JavaCast( "boolean", settingValue ) );
					cellStyle.setFont( font );
				break;
				case "leftborder":
					var borderStyle = cellStyle.getBorderLeftEnum()[ JavaCast( "string", UCase( settingValue ) ) ];
					cellStyle.setBorderLeft( borderStyle );
				break;
				case "leftbordercolor":
					cellStyle.setLeftBorderColor( getColor( workbook, settingValue ) );
				break;
				/*  TODO: Doesn't seem to do anything */
				case "locked":
					cellStyle.setLocked( JavaCast( "boolean", settingValue ) );
				break;
				/* TODO Implement when POI 3.16 available */
				/* case "quoteprefixed":
					cellStyle.setQuotePrefixed( JavaCast( "boolean", settingValue ) );
				break; */
				case "rightborder":
					var borderStyle = cellStyle.getBorderRightEnum()[ JavaCast( "string", UCase( settingValue ) ) ];
					cellStyle.setBorderRight( borderStyle );
				break;
				case "rightbordercolor":
					cellStyle.setRightBorderColor( getColor( workbook, settingValue ) );
				break;
				case "rotation":
					cellStyle.setRotation( JavaCast( "int", settingValue ) );
				break;
				case "strikeout":
					font = cloneFont( workbook,workbook.getFontAt( cellStyle.getFontIndex() ) );
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
					cellStyle.setTopBorderColor( getColor( workbook, settingValue ) );
				break;
				case "underline":
					font = cloneFont( workbook, workbook.getFontAt( cellStyle.getFontIndex() ) );
					font.setUnderline( JavaCast( "boolean", settingValue ) );
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
		var newFont = workbook.createFont();
		/*  copy the existing cell's font settings to the new font  */
		newFont.setBold( fontToClone.getBold() );
		newFont.setCharSet( fontToClone.getCharSet() );
		// xlsx fonts contain XSSFColor objects which may have been set as RGB
		newFont.setColor( isXmlFormat( workbook )? fontToClone.getXSSFColor(): fontToClone.getColor() );
		newFont.setFontHeight( fontToClone.getFontHeight() );
		newFont.setFontName( fontToClone.getFontName() );
		newFont.setItalic( fontToClone.getItalic() );
		newFont.setStrikeout( fontToClone.getStrikeout() );
		newFont.setTypeOffset( fontToClone.getTypeOffset() );
		newFont.setUnderline( fontToClone.getUnderline() );
		return newFont;
	}

	private numeric function getColorIndex( required string colorName ){
		var findColor = colorName.Trim().UCase();
		var indexedColors = loadPoi( "org.apache.poi.ss.usermodel.IndexedColors" );
		try{
			var color = indexedColors.valueOf( JavaCast( "string", findColor ) );
			return color.getIndex();
		}
		catch( any exception ){
			throw( type=exceptionType, message="Invalid Color", detail="The color provided (#colorName#) is not valid." );
		}
	}

	private any function getColor( required workbook, required string colorValue ){
		/* if colorValue is a preset name, returns the index */
		/* if colorValue is an RGB Triplet eg. "255,255,255" then the exact color object is returned for xlsx, or the nearest color's index if xls */
		var isRGB = ListLen( colorValue ) EQ 3;
		if( !isRGB )
			return getColorIndex( colorValue );
		var rgb = ListToArray( colorValue );
		if( isXmlFormat( workbook ) ){
			var javaColor = CreateObject( "Java", "java.awt.Color" ).init(
				JavaCast( "int", rgb[ 1 ] )
				,JavaCast( "int", rgb[ 2 ] )
				,JavaCast( "int", rgb[ 3 ] )
			);
			return loadPoi( "org.apache.poi.xssf.usermodel.XSSFColor" ).init( javaColor );
		}
		var palette = workbook.getCustomPalette();
		var similarExistingColor = palette.findSimilarColor(
			JavaCast( "int", rgb[ 1 ] )
			,JavaCast( "int", rgb[ 2 ] )
			,JavaCast( "int", rgb[ 3 ] )
		);
		return similarExistingColor.getIndex();
	}

	/* Override troublesome engine BIFs */

	private boolean function _isDate( required value ){
		if( !IsDate( value ) )
			return false;
		// Lucee will treat 01-23112 as a date!
		if( REFind( "\d\d[[:punct:]]\d{5,}", value ) ) // NB: member function doesn't work on dates in Lucee
			return false;
		return true;
	}

	/* ACF compatibility functions */
	private array function _QueryColumnArray( required query q ){
		try{
			return QueryColumnArray( q ); //Lucee
		}
		catch( any exception ){
			if( !exception.message CONTAINS "undefined" )
				rethrow;
			//ACF
			return q.ColumnList.ListToArray();
		}
	}

	private query function _QueryDeleteColumn( required query q, required string columnToDelete ){
		try{
			QueryDeleteColumn( q, columnToDelete ); //Lucee
			return q;
		}
		catch( any exception ){
			if( !exception.message CONTAINS "undefined" )
				rethrow;
			//ACF
			var columnMetaData = GetMetaData( q );
			var columns = [];
			var columnTypes = [];
			for( var column in columnMetaData ){
				if( column.name IS columnToDelete )
					continue;
				columns.Append( column.name );
				columnTypes.Append( column.typeName?: "VarChar" );
			}
			var data = [];
			for( row in q ){
				newRow = [];
				for( column in columns )
					newRow.Append( row[ column ] );
				data.Append( newRow );
			}
		}
		return _QueryNew( columns.ToList(), columnTypes.ToList(), data );
	}

	private query function _QueryNew( required string columnNameList, required string columnTypeList, required array data ){
		//ACF QueryNew() won't accept invalid variable names in the column name list, hence clunky workaround:
		//NB: 'data' should not contain structs since they use the column name as key: always use array of row arrays instead
		if( !isACF )
			return QueryNew( columnNameList, columnTypeList, data );
		var columnNames = columnNameList.ListToArray();
		var totalColumns = columnNames.Len();
		var tempColumnNames = [];
		var tempData = [];
		for( var i=1; i LTE totalColumns; i++ )
			tempColumnNames[ i ] = "column#i#";
		var q = QueryNew( tempColumnNames.ToList(), columnTypeList, data );
		// restore the real names without ACF barfing
		for( name in columnNames )
			name = JavaCast( "string", name );
		q.setColumnNames( columnNames );
		return q;
	}

}