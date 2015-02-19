component{

	variables.version = "0.2.0";

	variables.defaultFormats = { DATE = "m/d/yy", TIMESTAMP = "m/d/yy h:mm", TIME = "h:mm:ss" };
	variables.exceptionType	=	"cfsimplicity.lucee.spreadsheet";

	include "tools.cfm";
	include "formatting.cfm";

	function init(){
		return this;
	}

	void function flushPoiLoader(){
		lock scope="server" timeout="10"{
			StructDelete( server,"_poiLoader" );
		};
	}

	/* CUSTOM METHODS */

	any function workbookFromQuery( required query data,boolean addHeaderRow=true,boldHeaderRow=true,xmlformat=false ){
		var workbook = this.new( xmlformat=xmlformat );
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

	binary function binaryFromQuery( required query data,boolean addHeaderRow=true,boldHeaderRow=true,xmlformat=false ){
		/* Pass in a query and get a spreadsheet binary file ready to stream to the browser */
		var workbook = this.workbookFromQuery( argumentCollection=arguments );
		return this.readBinary( workbook );
	}

	void function downloadFileFromQuery(
		required query data
		,required string filename
		,boolean addHeaderRow=true
		,boolean boldHeaderRow=true
		,boolean xmlformat=false
		,string contentType
	){
		var safeFilename	=	this.filenameSafe( filename );
		var filenameWithoutExtension = safeFilename.REReplace( "\.xlsx?$","" );
		var binary = this.binaryFromQuery( data,addHeaderRow,boldHeaderRow,xmlformat );
		if( !arguments.KeyExists( "contentType" ) )
			arguments.contentType = xmlformat? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "application/msexcel";
		var extension = xmlFormat? "xlsx": "xls";
		header name="Content-Disposition" value="attachment; filename=#Chr(34)##filenameWithoutExtension#.#extension##Chr(34)#";
		content type=contentType variable="#binary#" reset="true";
	}

	void function writeFileFromQuery(
		required query data
		,required string filepath
		,boolean overwrite=false
		,boolean addHeaderRow=true
		,boldHeaderRow=true
		,xmlformat=false
	){
		if( !xmlformat AND ( ListLast( filepath,"." ) IS "xlsx" ) )
			arguments.xmlformat=true;
		var workbook = this.workbookFromQuery( data,addHeaderRow,boldHeaderRow,xmlFormat );
		if( xmlformat AND ( ListLast( filepath,"." ) IS "xls" ) )
			arguments.filePath &="x";// force to .xlsx
		this.write( workbook=workbook,filepath=filepath,overwrite=overwrite );
	}

	/* STANDARD CFML API */

	void function addColumn(
		required workbook
		,required string data /* Delimited list of cell values */
		,numeric startRow
		,numeric startColumn
		,boolean insert=true
		,string delimiter=","
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
						/* TODO: Handle other cell types ?  */
						cell = this.createCell( row,i );
						cell.setCellStyle( oldCell.getCellStyle() );
						cell.setCellValue( oldCell.getStringCellValue() );
						cell.setCellComment( oldCell.getCellComment() );
					}
				}
			}
			cell = this.createCell( row,cellNum );
			cell.setCellValue( JavaCast( "string",cellValue ) );
			rowNum++;
		}
	}

	void function addRow(
		required workbook
		,required string data /* Delimited list of data */
		,numeric row
		,numeric column=1
		,boolean insert=true
		,string delimiter=","
		,boolean handleEmbeddedCommas=true /* When true, values enclosed in single quotes are treated as a single element like in ACF. Only applies when the delimiter is a comma. */
	){
		if( arguments.KeyExists( "row" ) AND ( row LTE 0 ) )
			throw( type=exceptionType,message="Invalid row value",detail="The value for row must be greater than or equal to 1." );
		if( arguments.KeyExists( "column" ) AND ( column LTE 0 ) )
			throw( type=exceptionType,message="Invalid column value",detail="The value for column must be greater than or equal to 1." );
		if( !insert AND !arguments.KeyExists( "row") )
			throw( type=exceptionType,message="Missing row value",detail="To replace a row using 'insert', please specify the row to replace." );
		var lastRow = this.getNextEmptyRow( workbook );
		//If the requested row already exists ...
		if( arguments.KeyExists( "row" ) AND ( row LTE lastRow ) ){
			if( arguments.insert )
				shiftRows( workbook,row,lastRow,1 );//shift the existing rows down (by one row)
			else
				deleteRow( workbook,row );//otherwise, clear the entire row
		}
		var theRow = arguments.KeyExists( "row" )? this.createRow( workbook,arguments.row-1 ): this.createRow( workbook );
		var rowValues = this.parseRowData( data,delimiter,handleEmbeddedCommas );
		var cellNum = column - 1;
		var dateUtil = this.getDateUtil();
		for( var cellValue in rowValues ){
			cellValue=cellValue.Trim();
			var oldWidth = this.getActiveSheet( workbook ).getColumnWidth( cellNum );
			var cell = this.createCell( theRow,cellNum );
			var isDateColumn  = false;
			var dateMask  = "";
			if( IsNumeric( cellValue ) and !cellValue.REFind( "^0[\d]+" ) ){
				/*  NUMERIC  */
				/*  skip numeric strings with leading zeroes. treat those as text  */
				cell.setCellType( cell.CELL_TYPE_NUMERIC );
				cell.setCellValue( JavaCast( "double",cellValue ) );
			} else if( IsDate( cellValue ) ){
				/*  DATE  */
				cellFormat = this.getDateTimeValueFormat( cellValue );
				cell.setCellStyle( this.buildCellStyle( { workbook,dataFormat=cellFormat } ) );
				cell.setCellType( cell.CELL_TYPE_NUMERIC );
				/*  Excel's uses a different epoch than CF (1900-01-01 versus 1899-12-30). "Time" only values will not display properly without special handling - */
				if( cellFormat EQ variables.defaultFormats.TIME ){
					cellValue = TimeFormat( cellValue, "HH:MM:SS" );
				 	cell.setCellValue( dateUtil.convertTime( cellValue ) );
				} else {
					cell.setCellValue( ParseDateTime( cellValue ) );
				}
				dateMask = cellFormat;
				isDateColumn = true;
			} else if( cellValue.Len() ){
				/* STRING */
				cell.setCellType( cell.CELL_TYPE_STRING );
				cell.setCellValue( JavaCast( "string",cellValue ) );
			} else {
				/* EMPTY */
				cell.setCellType( cell.CELL_TYPE_BLANK );
				cell.setCellValue( "" );
			}
			this.autoSizeColumnFix( workbook,cellNum,isDateColumn,dateMask );
			cellNum++;
		}
	}

	void function addRows( required workbook,required query data,numeric row,numeric column=1,boolean insert=true ){
		var lastRow = this.getNextEmptyRow( workbook );
		if( arguments.KeyExists( "row" ) AND ( row LTE lastRow ) AND insert )
			shiftRows( workbook,row,lastRow,data.recordCount );
		var rowNum	=	arguments.keyExists( "row" )? row-1: this.getNextEmptyRow( workbook );
		var queryColumns = this.getQueryColumnFormats( workbook,data );
		var dateUtil = this.getDateUtil();
		var dateColumns  = {};
		for( var dataRow in data ){
			/* can't just call addRow() here since that function expects a comma-delimited list of data (probably not the greatest limitation ...) and the query data may have commas in it, so this is a bit redundant with the addRow() function */
			var theRow = this.createRow( workbook,rowNum,false );
			var cellNum = ( arguments.column-1 );
			/* Note: To properly apply date/number formatting:
   				- cell type must be CELL_TYPE_NUMERIC
   				- cell value must be applied as a java.util.Date or java.lang.Double (NOT as a string)
   				- cell style must have a dataFormat (datetime values only) */
   		/* populate all columns in the row */
   		for( var column in queryColumns ){
   			var cell 	= this.createCell( theRow, cellNum, false );
				var value = dataRow[ column.name ];
				var forceDefaultStyle = false;
				column.index = cellNum;

				/* Cast the values to the correct type, so data formatting is properly applied  */
				if( column.cellDataType IS "DOUBLE" AND IsNumeric( value ) ){
					cell.setCellValue( JavaCast("double", Val( value) ) );
				} else if( column.cellDataType IS "TIME" AND IsDate( value ) ){
					value = TimeFormat( ParseDateTime( value ),"HH:MM:SS");				
					cell.setCellValue( dateUtil.convertTime( value ) );
					forceDefaultStyle = true;
					var dateColumns[ column.name ] = { index=cellNum,type=column.cellDataType };
				} else if( column.cellDataType EQ "DATE" AND IsDate( value ) ){
					/* If the cell is NOT already formatted for dates, apply the default format 
					brand new cells have a styleIndex == 0  */
					var styleIndex = cell.getCellStyle().getDataFormat();
					var styleFormat = cell.getCellStyle().getDataFormatString();
					if( styleIndex EQ 0 OR NOT dateUtil.isADateFormat( styleIndex,styleFormat ) )
						forceDefaultStyle = true;
					cell.setCellValue( ParseDateTime( value ) );
					dateColumns[ column.name ] = { index=cellNum,type=column.cellDataType };
				} else if( column.cellDataType EQ "BOOLEAN" AND IsBoolean( value ) ){
					cell.setCellValue( JavaCast( "boolean",value ) );
				} else if( IsSimpleValue( value ) AND value.isEmpty() ){
					cell.setCellType( cell.CELL_TYPE_BLANK );
				} else {
					cell.setCellValue( JavaCast( "string",value ) );
				}
				/* Replace the existing styles with custom formatting  */
				if( column.KeyExists( "customCellStyle" ) ){
					cell.setCellStyle( column.customCellStyle );
					/* Replace the existing styles with default formatting (for readability). The reason we cannot just update the cell's style is because they are shared. So modifying it may impact more than just this one cell. */
				} else if( column.KeyExists( "defaultCellStyle" ) AND forceDefaultStyle ){
					cell.setCellStyle( column.defaultCellStyle );
				}
				cellNum++;
   		}
   		rowNum++;
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

	void function deleteSheet( required workbook,string sheetName,numeric sheetNumber ){
		this.validateSheetNameOrNumberWasProvided( argumentCollection=arguments );
		if( arguments.KeyExists( "sheetName" ) ){
			validateSheetName( sheetName );
			validateSheetExistsWithName( workbook,sheetName );
			arguments.sheetNumber = workbook.getSheetIndex( sheetName )+1;
		} else {
			validateSheetNumber( workbook,sheetNumber );
		}
		var sheetIndex = sheetNumber-1;
		this.deleteSheetAtIndex( workbook,sheetIndex );
	}

	void function deleteSheetNumber( required workbook,required numeric sheetNumber ){
		this.deleteSheet( argumentCollection=arguments );
	}

	void function formatCell( required workbook,required struct format,required numeric row,required numeric column,any cellStyle ){
		var cell = this.initializeCell( workbook,row,column );
		if( arguments.KeyExists( "cellStyle" ) )
			cell.setCellStyle( cellStyle );// reuse an existing style
		else
			cell.setCellStyle( this.buildCellStyle( workbook,format ) );
	}

	void function formatRow( required workbook,required struct format,required numeric rowNum ){
		var theRow = this.getActiveSheet( workbook ).getRow( arguments.rowNum-1 );
		if( IsNull( theRow ) )
			return;
		var cellIterator = theRow.cellIterator();
		while( cellIterator.hasNext() ){
			formatCell( workbook,format,rowNum,cellIterator.next().getColumnIndex()+1 );
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

	function new( string sheetName="Sheet1",boolean xmlformat=false ){
		var workbook = this.createWorkBook( sheetName,xmlFormat );
		this.createSheet( workbook,sheetName,xmlformat );
		setActiveSheet( workbook,sheetName );
		return workbook;
	}

	function read(
		required string src
		,string format
		,string columns //TODO
		,string columnNames //TODO
		,numeric headerRow
		,string rows //TODO
		,string sheetName
		,numeric sheetNumber // 1-based
		,boolean includeHeaderRow=false
		,boolean includeBlankRows=false
	){
		if( arguments.KeyExists( "query" ) )
			throw( type=exceptionType,message="Invalid argument 'query'.",details="Just use format='query' to return a query object" );
		if( arguments.KeyExists( "format" ) AND !ListFindNoCase( "query",format ) ) //,csv,html,tab,pipe
			throw( type=exceptionType,message="Invalid format",detail="Supported formats are: QUERY, HTML, CSV, TAB and PIPE" );
		if( arguments.KeyExists( "sheetName" ) AND arguments.KeyExists( "sheetNumber" ) )
			throw( type=exceptionType,message="Cannot provide both sheetNumber and sheetName arguments",detail="Only one of either 'sheetNumber' or 'sheetName' arguments may be provided." );
		 //TODO
		if( arguments.KeyExists( "columns" ) )
			throw( type=exceptionType,message="Argument not yet supported",detail="Sorry the 'columns' argument is not yet supported." );
		if( arguments.KeyExists( "columnNames" ) )
			throw( type=exceptionType,message="Argument not yet supported",detail="Sorry the 'columnNames' argument is not yet supported." );
		if( arguments.KeyExists( "rows" ) )
			throw( type=exceptionType,message="Argument not yet supported",detail="Sorry the 'rows' argument is not yet supported." );		
		//END TODO
		if( !FileExists( src ) )
			throw( type=exceptionType,message="Non-existent file",detail="Cannot find the file #src#." );
		var workbook = this.workbookFromFile( src );
		if( arguments.KeyExists( "sheetName" ) )
			this.setActiveSheet( workbook=workbook,sheetName=sheetName );
		if( !arguments.keyExists( "format" ) )
			return workbook;
		switch( format ){
			case "csv": case "tab": case "pipe":
				throw( type=exceptionType,message="Format not yet supported",detail="Sorry #format# is not yet supported as an ouput format" );
				break;
			case "html":
				throw( type=exceptionType,message="Format not yet supported",detail="Sorry #format# is not yet supported as an ouput format" );
				break;
			case "query":
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
				if( arguments.KeyExists( "includeBlankRows" ) )
					args.includeBlankRows=includeBlankRows;
				return this.sheetToQuery( argumentCollection=args );
		}
		return workbook;
	}

	binary function readBinary( required workbook ){
		var baos = CreateObject( "Java","org.apache.commons.io.output.ByteArrayOutputStream" ).init();
		workbook.write( baos );
		baos.flush();
		return baos.toByteArray();
	}

	void function removeSheet( required workbook,required string sheetName ){
		this.deleteSheet( argumentCollection=arguments );
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

	void function shiftRows( required workbook,required numeric startRow,numeric endRow=startRow,numeric offset=1 ){
		this.getActiveSheet( workbook ).shiftRows(
			JavaCast( "int",arguments.startRow - 1 )
			,JavaCast( "int",arguments.endRow - 1 )
			,JavaCast( "int",arguments.offset )
		);
	}

	void function write( required workbook,required string filepath,boolean overwrite=false,string password ){
		if( !overwrite AND FileExists( filepath ) )
			throw( type=exceptionType,message="File already exists",detail="The file path specified already exists. Use 'overwrite=true' if you wish to overwrite it." );
		// writeProtectWorkbook takes both a user name and a password, but since CF 9 tag only takes a password, just making up a user name 
		// TODO: workbook.isWriteProtected() returns true but the workbook opens without prompting for a password
		if( arguments.KeyExists( "password" ) AND !password.Trim().IsEmpty() )
			workbook.writeProtectWorkbook( JavaCast( "string",password ),JavaCast( "string","user" ) );
		var outputStream = CreateObject( "java","java.io.FileOutputStream" ).init( filepath );
		try{
			workbook.write( outputStream );
			outputStream.flush();
		}
		finally{
			// always close the stream. otherwise file may be left in a locked state if an unexpected error occurs
			outputStream.close();
		}
	}

	/* NOT YET IMPLEMENTED */

	private void function notYetImplemented(){
		throw( type=exceptionType,message="Function not yet implemented" );
	}

	function addFreezePane(){ notYetImplemented(); }
	function addImage(){ notYetImplemented(); }
	function addInfo(){ notYetImplemented(); }
	function addSplitPlane(){ notYetImplemented(); }
	function autoSizeColumn(){ notYetImplemented(); }
	function clearCellRange(){ notYetImplemented(); }
	function deleteColumn(){ notYetImplemented(); }
	function deleteColumns(){ notYetImplemented(); }
	function formatCellRange(){ notYetImplemented(); }
	function formatColumn(){ notYetImplemented(); }
	function formatColumns(){ notYetImplemented(); }
	function getCellComment(){ notYetImplemented(); }
	function getCellFormula(){ notYetImplemented(); }
	function getCellValue(){ notYetImplemented(); }
	function info(){ notYetImplemented(); }
	function mergeCells(){ notYetImplemented(); }
	function setCellComment(){ notYetImplemented(); }
	function setCellFormula(){ notYetImplemented(); }
	function setCellValue(){ notYetImplemented(); }
	function setColumnWidth(){ notYetImplemented(); }
	function setHeader(){ notYetImplemented(); }
	function setRowHeight(){ notYetImplemented(); }
	function shiftColumns(){ notYetImplemented(); }

}