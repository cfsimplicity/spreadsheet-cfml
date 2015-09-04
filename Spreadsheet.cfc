component{

	variables.version = "0.5.3";
	variables.poiLoaderName = "_poiLoader-" & Hash( GetCurrentTemplatePath() );

	variables.dateFormats = {
		DATE 				= "yyyy-mm-dd"
		,DATETIME		=	"yyyy-mm-dd HH:nn:ss"
		,TIME 			= "hh:mm:ss"
		,TIMESTAMP 	= "yyyy-mm-dd hh:mm:ss"
	};
	variables.exceptionType	=	"cfsimplicity.lucee.spreadsheet";

	include "tools.cfm";
	include "formatting.cfm";

	function init( struct dateFormats ){
		if( arguments.KeyExists( "dateFormats" ) )
			this.overrideDefaultDateFormats( arguments.dateFormats );
		return this;
	}

	private void function overrideDefaultDateFormats( required struct formats ){
		for( var format in formats ){
			if( !variables.dateFormats.KeyExists( format ) )
				throw( type=exceptionType,message="Invalid date format key",detail="'#format#' is not a valid dateformat key. Valid keys are DATE, DATETIME, TIME and TIMESTAMP" );
			variables.dateFormats[ format ]=formats[ format ];
		}
	}

	void function flushPoiLoader(){
		lock scope="server" timeout="10"{
			StructDelete( server,poiLoaderName );
		};
	}

	/* META INFO */
	public struct function getDateFormats(){
		return dateFormats;
	}

	/* CUSTOM METHODS */

	any function workbookFromQuery( required query data,boolean addHeaderRow=true,boldHeaderRow=true,xmlFormat=false ){
		var workbook = this.new( xmlFormat=xmlFormat );
		if( addHeaderRow ){
			var columns	=	QueryColumnArray( data );
			this.addRow( workbook,columns.ToList() );
			if( boldHeaderRow )
				this.formatRow( workbook,{ bold=true },1 );
			this.addRows( workbook,data,2,1 );
		} else {
			this.addRows( workbook,data );
		}
		return workbook;
	}

	binary function binaryFromQuery( required query data,boolean addHeaderRow=true,boldHeaderRow=true,xmlFormat=false ){
		/* Pass in a query and get a spreadsheet binary file ready to stream to the browser */
		var workbook = this.workbookFromQuery( argumentCollection=arguments );
		return this.readBinary( workbook );
	}

	void function downloadFileFromQuery(
		required query data
		,required string filename
		,boolean addHeaderRow=true
		,boolean boldHeaderRow=true
		,boolean xmlFormat=false
		,string contentType
	){
		var safeFilename	=	this.filenameSafe( filename );
		var filenameWithoutExtension = safeFilename.REReplace( "\.xlsx?$","" );
		var binary = this.binaryFromQuery( data,addHeaderRow,boldHeaderRow,xmlFormat );
		if( !arguments.KeyExists( "contentType" ) )
			arguments.contentType = xmlFormat? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "application/msexcel";
		var extension = xmlFormat? "xlsx": "xls";
		header name="Content-Disposition" value="attachment; filename=#Chr(34)##filenameWithoutExtension#.#extension##Chr(34)#";
		content type=contentType variable="#binary#" reset="true";
	}

	void function downloadCsvFromFile(
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
		arguments.format="csv";
		var csv=this.read( argumentCollection=arguments );
		var binary=ToBinary( ToBase64( csv.Trim() ) );
		var safeFilename	=	this.filenameSafe( filename );
		var filenameWithoutExtension = safeFilename.REReplace( "\.csv$","" );
		header name="Content-Disposition" value="attachment; filename=#Chr(34)##filenameWithoutExtension#.csv#Chr(34)#";
		content type=contentType variable="#binary#" reset="true";
	}

	void function writeFileFromQuery(
		required query data
		,required string filepath
		,boolean overwrite=false
		,boolean addHeaderRow=true
		,boldHeaderRow=true
		,xmlFormat=false
	){
		if( !xmlFormat AND ( ListLast( filepath,"." ) IS "xlsx" ) )
			arguments.xmlFormat=true;
		var workbook = this.workbookFromQuery( data,addHeaderRow,boldHeaderRow,xmlFormat );
		if( xmlFormat AND ( ListLast( filepath,"." ) IS "xls" ) )
			arguments.filePath &="x";// force to .xlsx
		this.write( workbook=workbook,filepath=filepath,overwrite=overwrite );
	}

	/* MAIN API */

	void function addColumn(
		required workbook
		,required string data /* Delimited list of cell values */
		,numeric startRow
		,numeric startColumn
		,boolean insert=true
		,string delimiter=","
		,boolean autoSize=false
	){
		var row 				= 0;
		var cell 				= 0;
		var oldCell 		= 0;
		var rowNum 			= ( arguments.KeyExists( "startRow" ) AND startRow )? startRow-1: 0;
		var cellNum 		= 0;
		var lastCellNum = 0;
		var cellValue 	= 0;
		if( arguments.KeyExists( "startColumn" ) ){
			cellNum = startColumn-1;
		} else {
			row = this.getActiveSheet( workbook ).getRow( rowNum );
			/* if this row exists, find the next empty cell number. note: getLastCellNum() returns the cell index PLUS ONE or -1 if not found */
			if( !IsNull( row ) AND row.getLastCellNum() GT 0 )
				cellNum = row.getLastCellNum();
			else
				cellNum = 0;
		}
		var columnNumber = cellNum+1;
		var columnData = ListToArray( data,delimiter );
		for( var cellValue in columnData ){
			/* if rowNum is greater than the last row of the sheet, need to create a new row  */
			if( rowNum GT this.getActiveSheet( workbook ).getLastRowNum() OR IsNull( this.getActiveSheet( workbook ).getRow( rowNum ) ) )
				row = this.createRow( workbook,rowNum );
			else
				row = this.getActiveSheet( workbook ).getRow( rowNum );
			/* POI doesn't have any 'shift column' functionality akin to shiftRows() so inserts get interesting */
			/* ** Note: row.getLastCellNum() returns the cell index PLUS ONE or -1 if not found */
			if( insert AND ( cellNum LT row.getLastCellNum() ) ){
				/*  need to get the last populated column number in the row, figure out which cells are impacted, and shift the impacted cells to the right to make room for the new data */
				lastCellNum = row.getLastCellNum();
				for( var i=lastCellNum; i EQ cellNum; i-- ){
					oldCell	=	row.getCell( JavaCast( "int",i-1 ) );
					if( !IsNull( oldCell ) ){
						cell = this.createCell( row,i );
						cell.setCellStyle( oldCell.getCellStyle() );
						var cellValue = this.getCellValueAsType( workbook,oldCell );
						this.setCellValueAsType( workbook,oldCell,cellValue );
						cell.setCellComment( oldCell.getCellComment() );
					}
				}
			}
			cell = this.createCell( row,cellNum );
			this.setCellValueAsType( workbook,cell,cellValue );
			rowNum++;
		}
		if( autoSize )
			this.autoSizeColumn( workbook,columnNumber );
	}

	void function addFreezePane(
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
			this.getActiveSheet( workbook ).createFreezePane( JavaCast( "int",freezeColumn ),JavaCast( "int",freezeRow ) );
			return;
		}
		// POI lets you specify an active pane if you use createSplitPane() here
		this.getActiveSheet( workbook ).createFreezePane(
			JavaCast( "int",freezeColumn )
			,JavaCast( "int",freezeRow )
			,JavaCast( "int",leftmostColumn )
			,JavaCast( "int",topRow )
		);
	}

	void function addImage(
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
			throw( type=exceptionType,message="Invalid argument combination",detail="You must provide either a file path or an image object" );
		if( arguments.KeyExists( "imageData" ) AND !arguments.KeyExists( "imageType" ) )
			throw( type=exceptionType,message="Invalid argument combination",detail="If you specify an image object, you must also provide the imageType argument" );
		var numberOfAnchorElements = ListLen( anchor );
		if( ( numberOfAnchorElements NEQ 4 ) AND ( numberOfAnchorElements NEQ 8 ) )
			throw( type=exceptionType,message="Invalid anchor argument",detail="The anchor argument must be a comma-delimited list of integers with either 4 or 8 elements" );
		//we'll need the image type int in all cases
		if( arguments.KeyExists( "filepath" ) ){
			if( !FileExists( filepath ) )
				throw( type=exceptionType,message="Non-existent file",detail="The specified file does not exist." );
			try{
				arguments.imageType = ListLast( FileGetMimeType( filepath ),"/" );
			}
			catch( any exception ){
				throw( type=exceptionType,message="Could Not Determine Image Type",detail="An image type could not be determined from the filepath provided" );
			}
		} else if( !arguments.KeyExists( "imageType" ) ){
			throw( type=exceptionType,message="Could Not Determine Image Type",detail="An image type could not be determined from the filepath or imagetype provided" );
		}
		arguments.imageType	=	imageType.UCase();
		switch( imageType ){
			case "DIB": case "EMF": case "JPEG": case "PICT": case "PNG": case "WMF":
				var imageTypeIndex = workbook[ "PICTURE_TYPE_" & imageType ];
			break;
			case "JPG":
				var imageTypeIndex = workbook.PICTURE_TYPE_JPEG;
			break;
			default:
				throw( type=exceptionType,message="Invalid Image Type",detail="Valid image types are DIB, EMF, JPG, JPEG, PICT, PNG, and WMF" );
		}
		if( arguments.KeyExists( "filepath" ) ){
			try{
				var inputStream = CreateObject( "java","java.io.FileInputStream" ).init( JavaCast("string",filepath ) );
				var ioUtils = this.loadPoi( "org.apache.poi.util.IOUtils" );
				var bytes = ioUtils.toByteArray( inputStream );
			}
			finally{
				inputStream.close();
			}
		} else {
			var bytes = ToBinary( imageData );
		}
		var imageIndex = workbook.addPicture( bytes,JavaCast( "int",imageTypeIndex ) );
		var theAnchor = this.loadPoi( "org.apache.poi.hssf.usermodel.HSSFClientAnchor" ).init();
		if( numberOfAnchorElements EQ 4 ){
			theAnchor.setRow1( JavaCast( "int",ListFirst( anchor )-1 ) );
			theAnchor.setCol1( JavaCast( "int",ListGetAt( anchor, 2 )-1 ) );
			theAnchor.setRow2( JavaCast( "int",ListGetAt( anchor, 3 )-1 ) );
			theAnchor.setCol2( JavaCast( "int",ListLast( anchor )-1 ) );
		} else if( numberOfAnchorElements EQ 8 ){
			theAnchor.setDx1( JavaCast( "int",ListFirst( anchor ) ) );
			theAnchor.setDy1( JavaCast( "int",ListGetAt( anchor,2 ) ) );
			theAnchor.setDx2( JavaCast( "int",ListGetAt( anchor,3 ) ) );
			theAnchor.setDy2( JavaCast( "int",ListGetAt( anchor,4 ) ) );
			theAnchor.setRow1( JavaCast( "int",ListGetAt( anchor,5 )-1 ) );
			theAnchor.setCol1( JavaCast( "int",ListGetAt( anchor,6 )-1 ) );
			theAnchor.setRow2( JavaCast( "int",ListGetAt( anchor,7 )-1 ) );
			theAnchor.setCol2( JavaCast( "int",ListLast( anchor )-1 ) );
		}
		/* TODO: need to look into createDrawingPatriarch() vs. getDrawingPatriarch() since create will kill any existing images. getDrawingPatriarch() throws  a null pointer exception when an attempt is made to add a second image to the spreadsheet  */
		var drawingPatriarch = getActiveSheet( workbook ).createDrawingPatriarch();
		var picture = drawingPatriarch.createPicture( theAnchor,imageIndex );
		/* Disabling this for now--maybe let people pass in a boolean indicating whether or not they want the image resized?
		 if this is a png or jpg, resize the picture to its original size (this doesn't work for formats other than jpg and png)
			<cfif imgTypeIndex eq getWorkbook().PICTURE_TYPE_JPEG or imgTypeIndex eq getWorkbook().PICTURE_TYPE_PNG>
				<cfset picture.resize() />
			</cfif>
		*/
	}

	void function addInfo( required workbook,required struct info ){
		/* Valid struct keys are author, category, lastauthor, comments, keywords, manager, company, subject, title */
		if( this.isBinaryFormat( workbook ) )
			this.addInfoBinary( workbook,info );
		else
			this.addInfoXml( workbook,info );
	}

	void function addRow(
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
			throw( type=exceptionType,message="Invalid row value",detail="The value for row must be greater than or equal to 1." );
		if( arguments.KeyExists( "column" ) AND ( column LTE 0 ) )
			throw( type=exceptionType,message="Invalid column value",detail="The value for column must be greater than or equal to 1." );
		if( !insert AND !arguments.KeyExists( "row") )
			throw( type=exceptionType,message="Missing row value",detail="To replace a row using 'insert', please specify the row to replace." );
		var lastRow = this.getNextEmptyRow( workbook );
		//If the requested row already exists...
		if( arguments.KeyExists( "row" ) AND ( row LTE lastRow ) ){
			if( arguments.insert )
				shiftRows( workbook,row,lastRow,1 );//shift the existing rows down (by one row)
			else
				deleteRow( workbook,row );//otherwise, clear the entire row
		}
		var theRow = arguments.KeyExists( "row" )? this.createRow( workbook,arguments.row-1 ): this.createRow( workbook );
		var rowValues = this.parseRowData( data,delimiter,handleEmbeddedCommas );
		var cellIndex = column-1;
		for( var cellValue in rowValues ){
			var cell = this.createCell( theRow,cellIndex );
			this.setCellValueAsType( workbook,cell,cellValue.Trim() )
			if( autoSizeColumns )
				this.autoSizeColumn( workbook,column );
			cellIndex++;
		}
	}

	void function addRows(
		required workbook
		,required query data
		,numeric row
		,numeric column=1
		,boolean insert=true
		,boolean autoSizeColumns=false
	){
		var lastRow = this.getNextEmptyRow( workbook );
		if( arguments.KeyExists( "row" ) AND ( row LTE lastRow ) AND insert )
			shiftRows( workbook,row,lastRow,data.recordCount );
		var rowNum	=	arguments.keyExists( "row" )? row-1: this.getNextEmptyRow( workbook );
		var queryColumns = this.getQueryColumnFormats( workbook,data );
		var dateUtil = this.getDateUtil();
		for( var dataRow in data ){
			var newRow = this.createRow( workbook,rowNum,false );
			var cellIndex = ( column-1 );
			/* Note: To properly apply date/number formatting:
 				- cell type must be CELL_TYPE_NUMERIC
 				- cell value must be applied as a java.util.Date or java.lang.Double (NOT as a string)
 				- cell style must have a dataFormat (datetime values only)
   		*/
   		/* populate all columns in the row */
   		for( var queryColumn in queryColumns ){
   			var cell 	= this.createCell( newRow, cellIndex, false );
				var value = dataRow[ queryColumn.name ];
				var forceDefaultStyle = false;
				queryColumn.index = cellIndex;
				/* Cast the values to the correct type, so data formatting is properly applied  */
				if( queryColumn.cellDataType IS "DOUBLE" AND IsNumeric( value ) ){
					cell.setCellValue( JavaCast( "double",Val( value ) ) );
				} else if( queryColumn.cellDataType IS "TIME" AND IsDate( value ) ){
					value = TimeFormat( ParseDateTime( value ),"HH:MM:SS");
					cell.setCellValue( dateUtil.convertTime( value ) );
					forceDefaultStyle = true;
				} else if( queryColumn.cellDataType EQ "DATE" AND IsDate( value ) ){
					/* If the cell is NOT already formatted for dates, apply the default format brand new cells have a styleIndex == 0  */
					var styleIndex = cell.getCellStyle().getDataFormat();
					var styleFormat = cell.getCellStyle().getDataFormatString();
					if( styleIndex EQ 0 OR NOT dateUtil.isADateFormat( styleIndex,styleFormat ) )
						forceDefaultStyle = true;
					cell.setCellValue( ParseDateTime( value ) );
				} else if( queryColumn.cellDataType EQ "BOOLEAN" AND IsBoolean( value ) ){
					cell.setCellValue( JavaCast( "boolean",value ) );
				} else if( IsSimpleValue( value ) AND !Len( value ) ){ //NB don't use member function: won't work if numeric
					cell.setCellType( cell.CELL_TYPE_BLANK );
				} else {
					cell.setCellValue( JavaCast( "string",value ) );
				}
				/* Replace the existing styles with custom formatting  */
				if( queryColumn.KeyExists( "customCellStyle" ) ){
					cell.setCellStyle( queryColumn.customCellStyle );
					/* Replace the existing styles with default formatting (for readability). The reason we cannot just update the cell's style is because they are shared. So modifying it may impact more than just this one cell. */
				} else if( queryColumn.KeyExists( "defaultCellStyle" ) AND forceDefaultStyle ){
					cell.setCellStyle( queryColumn.defaultCellStyle );
				}
				cellIndex++;
   		}
   		rowNum++;
		}
		if( autoSizeColumns ){
			var numberOfColumns = queryColumns.Len();
			var thisColumn = column;
			for( var i=thisColumn; i LTE numberOfColumns; i++ ){
				this.autoSizeColumn( workbook,thisColumn );
				thisColumn++;
			}
		}
	}

	void function addSplitPane(
		required workbook
		,required numeric xSplitPosition
		,required numeric ySplitPosition
		,required numeric leftmostColumn
		,required numeric topRow
		,string activePane="UPPER_LEFT" //Valid values are LOWER_LEFT, LOWER_RIGHT, UPPER_LEFT, and UPPER_RIGHT
	){
		var activeSheet = this.getActiveSheet( workbook );
		arguments.activePane = activeSheet[ "PANE_#activePane#" ];
		activeSheet.createSplitPane(
			JavaCast( "int",xSplitPosition )
			,JavaCast( "int",ySplitPosition )
			,JavaCast( "int",leftmostColumn )
			,JavaCast( "int",topRow )
			,JavaCast( "int",activePane )
		);
	}

	void function autoSizeColumn( required workbook,required numeric column,boolean useMergedCells=false ){
		if( column LTE 0 )
			throw( type=exceptionType,message="Invalid column value",detail="The value for column must be greater than or equal to 1." );
		/* Adjusts the width of the specified column to fit the contents. For performance reasons, this should normally be called only once per column. */
		var columnIndex = column-1;
		this.getActiveSheet( workbook ).autoSizeColumn( columnIndex,useMergedCells );
	}

	void function clearCell( required workbook,required numeric row,required numeric column ){
		/* Clears the specified cell of all styles and values */
		var defaultStyle  = workbook.getCellStyleAt( JavaCast( "short",0 ) );
		var rowIndex = row-1;
		var rowObject = this.getActiveSheet( workbook ).getRow( JavaCast( "int",rowIndex ) );
		if( IsNull( rowObject ) )
			return;
		var columnIndex = column-1;
		var cell = rowObject.getCell( JavaCast( "int",columnIndex ) );
		if( IsNull( cell ) )
			return;
		cell.setCellStyle( defaultStyle );
		cell.setCellType( cell.CELL_TYPE_BLANK );
	}

	void function clearCellRange(
		required workbook
		,required numeric startRow
		,required numeric startColumn
		,required numeric endRow
		,required numeric endColumn
	){
		/* Clears the specified cell range of all styles and values */
		for( var rowNumber=startRow; rowNumber LTE endRow; rowNumber++ ){
			for( var columnNumber=startColumn; columnNumber LTE endColumn; columnNumber++ ){
				clearCell( workbook,rowNumber,columnNumber );
			}
		}
	}

	void function createSheet( required workbook,string sheetName,overwrite=false ){
		if( arguments.KeyExists( "sheetName" ) )
			this.validateSheetName( sheetName );
		else
			arguments.sheetName = this.generateUniqueSheetName( workbook );
		if( !this.sheetExists( workbook=workbook,sheetName=sheetName ) ){
			workbook.createSheet( JavaCast( "String",sheetName ) );
			return;
		}
		/* sheet already exists with that name */
		if( !overwrite )
			throw( type=exceptionType,message="Sheet name already exists",detail="A sheet with the name '#sheetName#' already exists in this workbook" );
		/* OK to replace the existing */
		var sheetIndexToReplace = workbook.getSheetIndex( JavaCast( "string",sheetName) );
		this.deleteSheetAtIndex( workbook,sheetIndexToReplace );
		var newSheet = workbook.createSheet( JavaCast( "String",sheetName ) );
		var moveToIndex = sheetIndexToReplace;
		this.moveSheet( workbook,sheetName,moveToIndex );
	}

	void function deleteColumn( required workbook,required numeric column ){
		if( column LTE 0 )
			throw( type=exceptionType,message="Invalid column value",detail="The value for column must be greater than or equal to 1." );
			/* POI doesn't have remove column functionality, so iterate over all the rows and remove the column indicated */
		var rowIterator = this.getActiveSheet( workbook ).rowIterator();
		while( rowIterator.hasNext() ){
			var row = rowIterator.next();
			var cell = row.getCell( JavaCast( "int",column-1 ) );
			if( IsNull( cell ) )
				continue;
			row.removeCell( cell );
		}
	}

	void function deleteColumns( required workbook,required string range ){
		/* Validate and extract the ranges. Range is a comma-delimited list of ranges, and each value can be either a single number or a range of numbers with a hyphen. */
		var allRanges = this.extractRanges( range );
		for( var thisRange in allRanges ){
			if( thisRange.startAt EQ thisRange.endAt ){
				/* Just one row */
				this.deleteColumn( workbook,thisRange.startAt );
				continue;
			}
			for( var columnNumber=thisRange.startAt; columnNumber LTE thisRange.endAt; columnNumber++ ){
				this.deleteColumn( workbook,columnNumber );
			}
		}
	}

	void function deleteRow( required workbook,required numeric row ){
		/* Deletes the data from a row. Does not physically delete the row. */
		if( row LTE 0 )
			throw( type=exceptionType,message="Invalid row value",detail="The value for row must be greater than or equal to 1." );
		var rowToDelete = row-1;
		if( rowToDelete GTE this.getFirstRowNum( workbook ) AND rowToDelete LTE this.getLastRowNum( workbook ) ) //If this is a valid row, remove it
			this.getActiveSheet( workbook ).removeRow( this.getActiveSheet( workbook ).getRow( JavaCast( "int",rowToDelete ) ) );
	}

	void function deleteRows( required workbook,required string range ){
		/* Validate and extract the ranges. Range is a comma-delimited list of ranges, and each value can be either a single number or a range of numbers with a hyphen. */
		var allRanges = this.extractRanges( range );
		for( var thisRange in allRanges ){
			if( thisRange.startAt EQ thisRange.endAt ){
				/* Just one row */
				this.deleteRow( workbook,thisRange.startAt );
				continue;
			}
			for( var rowNumber=thisRange.startAt; rowNumber LTE thisRange.endAt; rowNumber++ ){
				this.deleteRow( workbook,rowNumber );
			}
		}
	}

	void function formatCell( required workbook,required struct format,required numeric row,required numeric column,any cellStyle ){
		var cell = this.initializeCell( workbook,row,column );
		if( arguments.KeyExists( "cellStyle" ) )
			cell.setCellStyle( cellStyle );// reuse an existing style
		else
			cell.setCellStyle( this.buildCellStyle( workbook,format ) );
	}

	void function formatCellRange(
		required workbook
		,required struct format
		,required numeric startRow
		,required numeric endRow
		,required numeric startColumn
		,required numeric endColumn
		){
		var style = this.buildCellStyle( workbook,format );
		for( var rowNumber=startRow; rowNumber LTE endRow; rowNumber++ ){
			for( var columnNumber=startColumn; columnNumber LTE endColumn; columnNumber++ ){
				this.formatCell( workbook,format,rowNumber,columnNumber,style );
			}
		}
	}

	void function formatColumn( required workbook,required struct format,required numeric column ){
		if( column LT 1 )
			throw( type=exceptionType,message="Invalid column value",detail="The column value must be greater than 0" );
		var rowIterator = this.getActiveSheet( workbook ).rowIterator();
		var columnNumber = column;
		while( rowIterator.hasNext() ){
			var rowNumber = rowIterator.next().getRowNum() + 1;
			this.formatCell( workbook,format,rowNumber,columnNumber );
		}
	}

	void function formatColumns( required workbook,required struct format,required string range ){
		/* Validate and extract the ranges. Range is a comma-delimited list of ranges, and each value can be either a single number or a range of numbers with a hyphen. */
		var allRanges = this.extractRanges( range );
		for( var thisRange in allRanges ){
			if( thisRange.startAt EQ thisRange.endAt ){
				/* Just one column */
				this.formatColumn( workbook,format,thisRange.startAt );
				continue;
			}
			for( var columnNumber=thisRange.startAt; columnNumber LTE thisRange.endAt; columnNumber++ ){
				this.formatColumn( workbook,format,columnNumber );
			}
		}
	}

	void function formatRow( required workbook,required struct format,required numeric row ){
		var rowIndex = row-1;
		var theRow = this.getActiveSheet( workbook ).getRow( rowIndex );
		if( IsNull( theRow ) )
			return;
		var cellIterator = theRow.cellIterator();
		while( cellIterator.hasNext() ){
			formatCell( workbook,format,row,cellIterator.next().getColumnIndex()+1 );
		}
	}

	void function formatRows( required workbook,required struct format,required string range ){
		/* Validate and extract the ranges. Range is a comma-delimited list of ranges, and each value can be either a single number or a range of numbers with a hyphen. */
		var allRanges = this.extractRanges( range );
		for( var thisRange in allRanges ){
			if( thisRange.startAt EQ thisRange.endAt ){
				/* Just one row */
				this.formatRow( workbook,format,thisRange.startAt );
				continue;
			}
			for( var rowNumber=thisRange.startAt; rowNumber LTE thisRange.endAt; rowNumber++ ){
				this.formatRow( workbook,format,rowNumber );
			}
		}
	}

	function getCellComment( required workbook,numeric row,numeric column ){
		if( arguments.keyExists( "row" ) AND !arguments.KeyExists( "column" ) )
			throw( type=exceptionType,message="Invalid argument combination",detail="If you specify the row you must also specify the column" );
		if( arguments.keyExists( "column" ) AND !arguments.KeyExists( "row" ) )
			throw( type=exceptionType,message="Invalid argument combination",detail="If you specify the column you must also specify the row" );
		if( arguments.KeyExists( "row" ) ){
			var cell = this.getCellAt( workbook,row,column );
			var commentObject = cell.getCellComment();
			if( !IsNull( commentObject ) ){
				return {
					author = commentObject.getAuthor()
					,comment = commentObject.getString().getString()
					,column = column
					,row = row
				}
			}
			return {};
		}
		/* TODO: Look into checking all sheets in the workbook */
		/* row and column weren't provided so loop over the whole sheet and return all the comments as an array of structs */
		var result = [];
		var rowIterator = this.getActiveSheet( workbook ).rowIterator();
		while( rowIterator.hasNext() ){
			var cellIterator = rowIterator.next().cellIterator();
			while( cellIterator.hasNext() ){
				var commentObject = cellIterator.next().getCellComment();
				if( !IsNull( commentObject ) ){
					var comment = {
						author = commentObject.getAuthor()
						,comment = commentObject.getString().getString()
						,column = column
						,row = row
					}
					comments.Append( comment );
				}
			}
		}
		return comments;
	}

	function getCellFormula( required workbook,numeric row,numeric column ){
		if( arguments.KeyExists( "row" ) AND arguments.KeyExists( "column" ) ){
			if( cellExists( workbook,row,column ) ){
				var cell = getCellAt( workbook,row,column );
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
					row = ( cell.getRowIndex() + 1 )
					,column = ( cell.getColumnIndex() + 1 )
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

	function getCellValue( required workbook,required numeric row,required numeric column ){
		if( !this.cellExists( workbook,row,column ) )
			return "";
		var rowIndex = row-1;
		var columnIndex = column-1;
		var rowObject = this.getActiveSheet( workbook ).getRow( JavaCast( "int",rowIndex ) );
		var cell = rowObject.getCell( JavaCast( "int",columnIndex ) );
		var formatter = this.getFormatter();
		if( cell.getCellType() EQ cell.CELL_TYPE_FORMULA ){
			var formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
			return formatter.formatCellValue( cell,formulaEvaluator );
		}
		return formatter.formatCellValue( cell );
	}

	struct function info( required workbook ){
		/*
		workbook properties returned in the struct are:
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
		 //format specific metadata
		var info = this.isBinaryFormat( workbook )? this.binaryInfo( workbook ): this.xmlInfo( workbook );
		//common properties
		info.sheets = workbook.getNumberOfSheets();
		var sheetnames = [];
		if( IsNumeric( info.sheets ) ){
			for( var i=1; i LTE info.sheets; i++ ){
				sheetnames.Append( workbook.getSheetName( JavaCast( "int",i-1 ) ) );
			}
			info.sheetnames = sheetnames.ToList();
		}
		info.spreadSheetType = this.isXmlFormat( workbook )? "Excel (2007)": "Excel";
		return info;
	}

	void function hideColumn( required workbook,required numeric column ){
		this.toggleColumnHidden( workbook,column,true );
	}

	boolean function isBinaryFormat( required workbook ){
		return workbook.getClass().getCanonicalName() IS "org.apache.poi.hssf.usermodel.HSSFWorkbook";
	}

	boolean function isXmlFormat( required workbook ){
		return workbook.getClass().getCanonicalName() IS "org.apache.poi.xssf.usermodel.XSSFWorkbook";
	}

	void function mergeCells(
		required workbook
		,required numeric startRow
		,required numeric endRow
		,required numeric startColumn
		,required numeric endColumn
		,boolean emptyInvisibleCells=false
	){
		if( startRow LT 1 OR startRow GT endRow )
			throw( type=exceptionType,message="Invalid startRow or endRow",detail="Row values must be greater than 0 and the startRow cannot be greater than the endRow." );
		if( startColumn LT 1 OR startColumn GT endColumn )
			throw( type=exceptionType,message="Invalid startColumn or endColumn",detail="Column values must be greater than 0 and the startColumn cannot be greater than the endColumn." );
		var cellRangeAddress = loadPoi( "org.apache.poi.ss.util.CellRangeAddress" ).init(
			JavaCast( "int",startRow - 1 )
			,JavaCast( "int",endRow - 1 )
			,JavaCast( "int",startColumn - 1 )
			,JavaCast( "int",endColumn - 1 )
		);
		this.getActiveSheet( workbook ).addMergedRegion( cellRangeAddress );
		if( !emptyInvisibleCells )
			return;
		// stash the value to retain
		var visibleValue	=	getCellValue( workbook,startRow,startColumn );
		//empty all cells in the merged region
		setCellRangeValue( workbook,"",startRow,endRow,startColumn,endColumn );
		//restore the stashed value
		setCellValue( workbook,visibleValue,startRow,startColumn );
	}

	function new( string sheetName="Sheet1",boolean xmlformat=false ){
		var workbook = this.createWorkBook( sheetName,xmlFormat );
		this.createSheet( workbook,sheetName,xmlformat );
		setActiveSheet( workbook,sheetName );
		return workbook;
	}

	function read(
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
	){
		if( arguments.KeyExists( "query" ) )
			throw( type=exceptionType,message="Invalid argument 'query'.",details="Just use format='query' to return a query object" );
		if( arguments.KeyExists( "format" ) AND !ListFindNoCase( "query,html,csv",format ) )
			throw( type=exceptionType,message="Invalid format",detail="Supported formats are: 'query', 'html' and 'csv'" );
		if( arguments.KeyExists( "sheetName" ) AND arguments.KeyExists( "sheetNumber" ) )
			throw( type=exceptionType,message="Cannot provide both sheetNumber and sheetName arguments",detail="Only one of either 'sheetNumber' or 'sheetName' arguments may be provided." );
		if( !FileExists( src ) )
			throw( type=exceptionType,message="Non-existent file",detail="Cannot find the file #src#." );
		var workbook = this.workbookFromFile( src );
		if( arguments.KeyExists( "sheetName" ) )
			this.setActiveSheet( workbook=workbook,sheetName=sheetName );
		if( !arguments.keyExists( "format" ) )
			return workbook;
		var args = {
			workbook = workbook
		};
		if( arguments.KeyExists( "sheetName" ) )
			args.sheetName = sheetName;
		if( arguments.KeyExists( "sheetNumber" ) )
			args.sheetNumber = sheetNumber;
		if( arguments.KeyExists( "headerRow" ) ){
			args.headerRow=headerRow;
			args.includeHeaderRow = includeHeaderRow;
		}
		if( arguments.KeyExists( "rows" ) )
			args.rows = rows;
		if( arguments.KeyExists( "columns" ) )
			args.columns = columns;
		if( arguments.KeyExists( "columnNames" ) )
			args.columnNames = columnNames;
		args.includeBlankRows=includeBlankRows;
		args.fillMergedCellsWithVisibleValue=fillMergedCellsWithVisibleValue;
		args.includeHiddenColumns=includeHiddenColumns;
		args.includeRichTextFormatting=includeRichTextFormatting;
		var generatedQuery=this.sheetToQuery( argumentCollection=args );
		if( format IS "query" )
			return generatedQuery;
		var args={
			query=generatedQuery
		};
		if( arguments.KeyExists( "headerRow" ) ){
			args.headerRow=headerRow;
			args.includeHeaderRow = includeHeaderRow;
		}
		switch( format ){
			case "csv": return this.queryToCsv( argumentCollection=args );
			case "html": return this.queryToHtml( argumentCollection=args );
		}
	}

	binary function readBinary( required workbook ){
		var baos = CreateObject( "Java","org.apache.commons.io.output.ByteArrayOutputStream" ).init();
		workbook.write( baos );
		baos.flush();
		return baos.toByteArray();
	}

	void function removeSheet( required workbook,required string sheetName ){
		validateSheetName( sheetName );
		validateSheetExistsWithName( workbook,sheetName );
		arguments.sheetNumber = workbook.getSheetIndex( sheetName )+1;
		var sheetIndex = sheetNumber-1;
		this.deleteSheetAtIndex( workbook,sheetIndex );
	}

	void function removeSheetNumber( required workbook,required numeric sheetNumber ){
		validateSheetNumber( workbook,sheetNumber );
		var sheetIndex = sheetNumber-1;
		this.deleteSheetAtIndex( workbook,sheetIndex );
	}

	void function renameSheet( required workbook,required string sheetName,required numeric sheetNumber ){
		this.validateSheetName( sheetName );
		this.validateSheetNumber( workbook,sheetNumber );
		var sheetIndex = sheetNumber-1;
		var foundAt = workbook.getSheetIndex( JavaCast( "string",sheetName ) );
		if( ( foundAt GT 0 ) AND ( foundAt NEQ sheetIndex ) )
			throw( type=exceptionType,message="Invalid Sheet Name [#sheetName#]",detail="The workbook already contains a sheet named [#sheetName#]. Sheet names must be unique" );
		workbook.setSheetName( JavaCast( "int",sheetIndex ),JavaCast( "string",sheetName ) );
	}

	void function setActiveSheet( required workbook,string sheetName,numeric sheetNumber ){
		this.validateSheetNameOrNumberWasProvided( argumentCollection=arguments );
		if( arguments.KeyExists( "sheetName" ) ){
			this.validateSheetExistsWithName( workbook,sheetName );
			sheetNumber = workbook.getSheetIndex( JavaCast( "string",sheetName ) ) + 1;
		}
		this.validateSheetNumber( workbook,sheetNumber )
		workbook.setActiveSheet( JavaCast( "int",sheetNumber - 1 ) );
	}

	void function setActiveSheetNumber( required workbook,numeric sheetNumber ){
		this.setActiveSheet( argumentCollection=arguments );
	}

	void function setCellComment(
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
		var drawingPatriarch = this.getActiveSheet( workbook ).createDrawingPatriarch();
		var commentString = this.loadPoi( "org.apache.poi.hssf.usermodel.HSSFRichTextString" ).init( JavaCast( "string",comment.comment ) );
		var javaColorRGB = 0;
		if( comment.KeyExists( "anchor" ) )
			var clientAnchor = loadPoi( "org.apache.poi.hssf.usermodel.HSSFClientAnchor" ).init(
				JavaCast( "int",0 )
				,JavaCast( "int",0 )
				,JavaCast( "int",0 )
				,JavaCast( "int",0 )
				,JavaCast( "int",ListGetAt( comment.anchor,1 ) )
				,JavaCast( "int",ListGetAt( comment.anchor,2 ) )
				,JavaCast( "int",ListGetAt( comment.anchor,3 ) )
				,JavaCast( "int",ListGetAt( comment.anchor,4 ) )
			);
		else
			var clientAnchor = loadPoi( "org.apache.poi.hssf.usermodel.HSSFClientAnchor" ).init(
				JavaCast( "int",0 )
				,JavaCast( "int",0 )
				,JavaCast( "int",0 )
				,JavaCast( "int",0 )
				,JavaCast( "int",column )
				,JavaCast( "int",row )
				,JavaCast( "int",column+2 )
				,JavaCast( "int",row+2 )
			);
		var commentObject = drawingPatriarch.createComment( clientAnchor );
		if( comment.KeyExists( "author" ) )
			commentObject.setAuthor( JavaCast( "string",comment.author ) );
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
			if( comment.KeyExists( "bold" ) ){
				if( comment.bold )
					font.setBoldWeight( font.BOLDWEIGHT_BOLD );
				else
					font.setBoldWeight( font.BOLDWEIGHT_NORMAL );
			}
			if( comment.KeyExists( "color" ) )
				font.setColor( JavaCast( "int",getColorIndex( comment.color ) ) );
			if( comment.KeyExists( "font" ) )
				font.setFontName( JavaCast( "string",comment.font ) );
			if( comment.KeyExists( "italic" ) )
				font.setItalic( JavaCast( "string",comment.italic ) );
			if( comment.KeyExists( "size" ) )
				font.setFontHeightInPoints( JavaCast( "int",comment.size ) );
			if( comment.KeyExists( "strikeout" ) )
				font.setStrikeout( JavaCast( "boolean",comment.strikeout ) );
			if( comment.KeyExists( "underline" ) )
				font.setUnderline( JavaCast( "boolean",comment.underline ) );
			commentString.applyFont( font );
		}
		if( comment.KeyExists( "fillColor" ) ){
			javaColorRGB = this.getJavaColorRGB( comment.fillColor );
			commentObject.setFillColor(
				JavaCast( "int",javaColorRGB.red )
				,JavaCast( "int",javaColorRGB.green )
				,JavaCast( "int",javaColorRGB.blue )
			);
		}
		/*
			Horizontal alignment can be left, center, right, justify, or distributed. Note that the constants on the Java class are slightly different in some cases:
			'center' = CENTERED
			'justify' = JUSTIFIED
		 */
		if( comment.KeyExists( "horizontalAlignment" ) ){
			if( comment.horizontalAlignment.UCase() IS "CENTER" )
				comment.horizontalAlignment = "CENTERED";
			if( comment.horizontalAlignment.UCase() IS "JUSTIFY" )
				comment.horizontalAlignment = "JUSTIFIED";
			commentObject.setHorizontalAlignment( JavaCast( "int",commentObject[ "HORIZONTAL_ALIGNMENT_" & comment.horizontalalignment.UCase() ] ) );
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
		 	commentObject.setLineStyle( JavaCast( "int",commentObject[ "LINESTYLE_" & comment.lineStyle.UCase() ] ) );
		if( comment.KeyExists( "lineStyleColor" ) ){
			javaColorRGB = this.getJavaColorRGB( comment.lineStyleColor );
			commentObject.setLineStyleColor(
				JavaCast( "int",javaColorRGB.red )
				,JavaCast( "int",javaColorRGB.green )
				,JavaCast( "int",javaColorRGB.blue )
			);
		}
		/* Vertical alignment can be top, center, bottom, justify, and distributed. Note that center and justify are DIFFERENT than the constants for horizontal alignment, which are CENTERED and JUSTIFIED. */
		if( comment.KeyExists( "verticalAlignment" ) )
			commentObject.setVerticalAlignment( JavaCast( "int",commentObject[ "VERTICAL_ALIGNMENT_" & comment.verticalAlignment.UCase() ] ) );
		if( comment.KeyExists( "visible" ) )
			commentObject.setVisible( JavaCast( "boolean",comment.visible ) );//doesn't seem to work
		commentObject.setString( commentString );
		var cell = this.initializeCell( workbook,row,column );
		cell.setCellComment( commentObject );
	}

	void function setCellFormula( required workbook,required string formula,required numeric row,required numeric column ){
		//Automatically create the cell if it does not exist, instead of throwing an error
		var cell = initializeCell( workbook,row,column );
		cell.setCellFormula( JavaCast( "string",formula ) );
	}

	void function setCellValue( required workbook,required value,required numeric row,required numeric column ){
		//Automatically create the cell if it does not exist, instead of throwing an error
		var cell = initializeCell( workbook,row,column );
		this.setCellValueAsType( workbook,cell,value );
	}

	void function setCellRangeValue(
		required workbook
		,required value
		,required numeric startRow
		,required numeric endRow
		,required numeric startColumn
		,required numeric endColumn
	){
		/* Sets the same value to a range of cells */
		for( var rowNumber=startRow; rowNumber LTE endRow; rowNumber++ ){
			for( var columnNumber=startColumn; columnNumber LTE endColumn; columnNumber++ ){
				setCellValue( workbook,value,rowNumber,columnNumber );
			}
		}
	}

	void function setColumnWidth( required workbook,required numeric column,required numeric width ){
		var columnIndex = column-1;
		this.getActiveSheet( workbook ).setColumnWidth( JavaCast( "int",columnIndex ),JavaCast( "int",width*256 ) );
	}

	void function setFooter(
		required workbook
		,string leftFooter=""
		,string centerFooter=""
		,string rightFooter=""
	){
		if( !centerFooter.isEmpty() )
			this.getActiveSheet( workbook ).getFooter().setCenter( JavaCast( "string",centerFooter ) );
		if( !leftFooter.isEmpty() )
			this.getActiveSheet( workbook ).getFooter().setleft( JavaCast( "string",leftFooter ) );
		if( !rightFooter.isEmpty() )
			this.getActiveSheet( workbook ).getFooter().setright( JavaCast( "string",rightFooter ) );
	}

	void function setHeader(
		required workbook
		,string leftHeader=""
		,string centerHeader=""
		,string rightHeader=""
	){
		if( !centerHeader.isEmpty() )
			this.getActiveSheet( workbook ).getHeader().setCenter( JavaCast( "string",centerHeader ) );
		if( !leftHeader.isEmpty() )
			this.getActiveSheet( workbook ).getHeader().setleft( JavaCast( "string",leftHeader ) );
		if( !rightHeader.isEmpty() )
			this.getActiveSheet( workbook ).getHeader().setright( JavaCast( "string",rightHeader ) );
	}

	void function setRowHeight( required workbook,required numeric row,required numeric height ){
		var rowIndex = row-1;
		this.getActiveSheet( workbook ).getRow( JavaCast( "int",rowIndex ) ).setHeightInPoints( JavaCast( "int",height ) );
	}

	void function shiftColumns( required workbook,required numeric start,numeric end=start,numeric offset=1 ){
		if( start LTE 0 )
			throw( type=exceptionType,message="Invalid start value",detail="The start value must be greater than or equal to 1" );
		if( arguments.KeyExists( "end" ) AND ( end LTE 0 OR end LT start ) )
			throw( type=exceptionType,message="Invalid end value",detail="The end value must be greater than or equal to the start value" );
		var rowIterator = this.getActiveSheet( workbook ).rowIterator();
		var startIndex = start-1;
		var endIndex = arguments.KeyExists( "end" )? end-1: startIndex;
		while( rowIterator.hasNext() ){
			var row = rowIterator.next();
			if( offset GT 0 ){
				for( var i=endIndex; i GTE startIndex; i-- ){
					var tempCell = row.getCell( JavaCast( "int",i ) );
					var cell = this.createCell( row,i+offset );
					if( !IsNull( tempCell ) ){
						this.setCellValueAsType( workbook,cell,this.getCellValueAsType( workbook,tempCell ) );
						cell.setCellStyle( tempCell.getCellStyle() );
					}
				}
			} else {
				for( var i=startIndex; i LTE endIndex; i++ ){
					var tempCell = row.getCell( JavaCast( "int",i ) );
					var cell = createCell( row,i+offset );
					if( !IsNull( tempCell ) ){
						this.setCellValueAsType( workbook,cell,this.getCellValueAsType( workbook,tempCell ) );
						cell.setCellStyle( tempCell.getCellStyle() );
					}
				}
			}
		}
		// clean up any columns that need to be deleted after the shift
		var numberColsShifted = ( endIndex-startIndex )+1;
		var numberColsToDelete = Abs( offset );
		if( numberColsToDelete GT numberColsShifted )
			numberColsToDelete = numberColsShifted;
		if( offset GT 0 ){
			var stopValue = ( startIndex + numberColsToDelete )-1;
			for( var i=startIndex; i LTE stopValue; i++ ){
				this.deleteColumn( workbook,i+1 );
			}
			return;
		}
		var stopValue = ( endIndex - numberColsToDelete )+1;
		for( var i=endIndex; i GTE stopValue; i-- ){
			this.deleteColumn( workbook,i+1 );
		}
	}

	void function shiftRows( required workbook,required numeric start,numeric end=start,numeric offset=1 ){
		this.getActiveSheet( workbook ).shiftRows(
			JavaCast( "int",arguments.start - 1 )
			,JavaCast( "int",arguments.end - 1 )
			,JavaCast( "int",arguments.offset )
		);
	}

	void function showColumn( required workbook,required numeric column ){
		this.toggleColumnHidden( workbook,column,false );
	}

	void function write( required workbook,required string filepath,boolean overwrite=false,string password ){
		if( !overwrite AND FileExists( filepath ) )
			throw( type=exceptionType,message="File already exists",detail="The file path specified already exists. Use 'overwrite=true' if you wish to overwrite it." );
		// writeProtectWorkbook takes both a user name and a password, but since CF 9 tag only takes a password, just making up a user name
		// TODO: workbook.isWriteProtected() returns true but the workbook opens without prompting for a password
		if( arguments.KeyExists( "password" ) AND !password.Trim().IsEmpty() )
			workbook.writeProtectWorkbook( JavaCast( "string",password ),JavaCast( "string","user" ) );
		lock name="#filepath#" timeout=5{
			var outputStream = CreateObject( "java","java.io.FileOutputStream" ).init( filepath );
		}
		try{
			lock name="#filepath#" timeout=5{
				workbook.write( outputStream );
			}
			outputStream.flush();
		}
		finally{
			// always close the stream. otherwise file may be left in a locked state if an unexpected error occurs
			outputStream.close();
		}
	}

}
