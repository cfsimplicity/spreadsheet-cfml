component{

	variables.defaultFormats = { DATE = "m/d/yy", TIMESTAMP = "m/d/yy h:mm", TIME = "h:mm:ss" };
	variables.exceptionType	=	"cfsimplicity.Railo.Spreadsheet";

	function init(){
		variables.formatting = New formatting( exceptionType );
		variables.tools = New tools( formatting,defaultFormats,exceptionType );
		return this;
	}

	/* CUSTOM METHODS */

	binary function binaryFromQuery( required query data,boolean addHeaderRow=true,boldHeaderRow=true,xmlformat=false ){
		/* Pass in a query and get a spreadsheet binary file ready to stream to the browser */
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
		return this.readBinary( workbook );
	}

	void function downloadFileFromQuery(
		required query data
		,required string filename
		,boolean addHeaderRow=true
		,boldHeaderRow=true
		,xmlformat=false
		,contentType
	){
		var safeFilename	=	tools.filenameSafe( filename );
		var filenameWithoutExtension = safeFilename.REReplace( "\.xlsx?$","" );
		var binary = binaryFromQuery( data,addHeaderRow,boldHeaderRow,xmlformat );
		if( !arguments.KeyExists( "contentType" ) )
			arguments.contentType = xmlformat? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "application/msexcel";
		var extension = xmlFormat? "xlsx": "xls";
		header name="Content-Disposition" value="attachment; filename=#Chr(34)##filenameWithoutExtension#.#extension##Chr(34)#";
		content type=contentType variable="#binary#" reset="true";
	}

	/* STANDARD CFML API */

	void function addColumn(
		required workbook
		,required string data /* Delimited list of cell values */
		,numeric startRow
		,numeric column
		,boolean insert=true
		,string delimiter=","
	){
		var row 				= 0;
		var cell 				= 0;
		var oldCell 		= 0;
		var rowNum 			= 0;
		var cellNum 		= 0;
		var lastCellNum = 0;
		var cellValue 	= 0;
		if( arguments.KeyExists( "startRow" ) )
			rowNum = startRow-1;
		if( arguments.KeyExists( "column" ) ){
			cellNum = column-1;
		} else {
			row = tools.getActiveSheet( workbook ).getRow( rowNum );
			/* if this row exists, find the next empty cell number. note: getLastCellNum() 
				returns the cell index PLUS ONE or -1 if not found */
			if( !IsNull( row ) AND row.getLastCellNum() GT 0 )
				cellNum = row.getLastCellNum();
			else
				cellNum = 0;
		}
		var columnData = ListToArray( data,delimiter );
		for( var cellValue in columnData ){
			/* if rowNum is greater than the last row of the sheet, need to create a new row  */
			if( rowNum GT tools.getActiveSheet( workbook ).getLastRowNum() OR IsNull( tools.getActiveSheet( workbook ).getRow( rowNum ) ) )
				row = tools.createRow( workbook,rowNum );
			else
				row = tools.getActiveSheet( workbook ).getRow( rowNum );
			/* POI doesn't have any 'shift column' functionality akin to shiftRows() so inserts get interesting */
			/* ** Note: row.getLastCellNum() returns the cell index PLUS ONE or -1 if not found */
			if( insert AND cellNum LT row.getLastCellNum() ){
				/*  need to get the last populated column number in the row, figure out which 
						cells are impacted, and shift the impacted cells to the right to make 
						room for the new data */
				lastCellNum = row.getLastCellNum();
				for( var i=lastCellNum; i EQ cellNum; i-- ){
					oldCell	=	row.getCell( JavaCast( "int",i-1 ) );
					if( !IsNull( oldCell ) ){
						/* TODO: Handle other cell types ?  */
						cell = tools.createCell( row,i );
						cell.setCellStyle( oldCell.getCellStyle() );
						cell.setCellValue( oldCell.getStringCellValue() );
						cell.setCellComment( oldCell.getCellComment() );
					}
				}
			}
			cell = tools.createCell( row,cellNum );
			cell.setCellValue( JavaCast( "string",cellValue ) );
			rowNum++;
		}
	}

	void function addRow(
		required workbook
		,required string data /* Delimited list of data */
		,numeric startRow
		,numeric startColumn=1
		,boolean insert=true
		,string delimiter=","
		,boolean handleEmbeddedCommas=true /* When true, values enclosed in single quotes are treated as a single element like in ACF. Only applies when the delimiter is a comma. */
	){
		var lastRow = tools.getNextEmptyRow( workbook );
		//If the requested row already exists ...
		if( arguments.KeyExists( "startRow" ) AND startRow LTE lastRow ){
			shiftRows( startRow,lastRow,1 );//shift the existing rows down (by one row)
		else
			deleteRow( startRow );//otherwise, clear the entire row
		}
		var theRow = arguments.KeyExists( "startRow" )? tools.createRow( workbook,arguments.startRow-1 ): tools.createRow( workbook );
		var rowValues = tools.parseRowData( data,delimiter,handleEmbeddedCommas );
		var cellNum = startColumn - 1;
		var dateUtil = tools.getDateUtil();
		for( var cellValue in rowValues ){
			cellValue=cellValue.Trim();
			var oldWidth = tools.getActiveSheet( workbook ).getColumnWidth( cellNum );
			var cell = tools.createCell( theRow,cellNum );
			var isDateColumn  = false;
			var dateMask  = "";
			if( IsNumeric( cellValue ) and !cellValue.REFind( "^0[\d]+" ) ){
				/*  NUMERIC  */
				/*  skip numeric strings with leading zeroes. treat those as text  */
				cell.setCellType( cell.CELL_TYPE_NUMERIC );
				cell.setCellValue( JavaCast( "double",cellValue ) );
			} else if( IsDate( cellValue ) ){
				/*  DATE  */
				cellFormat = tools.getDateTimeValueFormat( cellValue );
				cell.setCellStyle( formatting.buildCellStyle( { workbook,dataFormat=cellFormat } ) );
				cell.setCellType( cell.CELL_TYPE_NUMERIC );
				/*  Excel's uses a different epoch than CF (1900-01-01 versus 1899-12-30). "Time" 
				only values will not display properly without special handling - */
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
			tools.autoSizeColumnFix( workbook,cellNum,isDateColumn,dateMask );
			cellNum++;
		}
	}

	void function addRows( required workbook,required query data,numeric row,numeric column=1,boolean insert=true ){
		var lastRow = tools.getNextEmptyRow( workbook );
		if( arguments.KeyExists( "row" ) AND row LTE lastRow AND insert )
			shiftRows( row,lastRow,data.recordCount );
		var rowNum	=	arguments.keyExists( "row" )? row-1: tools.getNextEmptyRow( workbook );
		var queryColumns = tools.getQueryColumnFormats( workbook,data );
		var dateUtil = tools.getDateUtil();
		var dateColumns  = {};
		for( var dataRow in data ){
			/* can't just call addRow() here since that function expects a comma-delimited 
					list of data (probably not the greatest limitation ...) and the query 
					data may have commas in it, so this is a bit redundant with the addRow() 
					function */
			var theRow = tools.createRow( workbook,rowNum,false );
			var cellNum = ( arguments.column-1 );
			/* Note: To properly apply date/number formatting:
   				- cell type must be CELL_TYPE_NUMERIC
   				- cell value must be applied as a java.util.Date or java.lang.Double (NOT as a string)
   				- cell style must have a dataFormat (datetime values only) */
   		/* populate all columns in the row */
   		for( var column in queryColumns ){
   			var cell 	= tools.createCell( theRow, cellNum, false );
				var value = dataRow[ column.name ];
				var forceDefaultStyle = false;
				column.index = cellNum;

				/* Cast the values to the correct type, so data formatting is properly applied  */
				if( column.cellDataType IS "DOUBLE" AND isNumeric( value ) ){
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
					/* Replace the existing styles with default formatting (for readability). The reason we cannot 
					just update the cell's style is because they are shared. So modifying it may impact more than 
					just this one cell. */
				} else if( column.KeyExists( "defaultCellStyle" ) AND forceDefaultStyle ){
					cell.setCellStyle( column.defaultCellStyle );
				}
				cellNum++;
   		}
   		rowNum++;
		}
	}

	void function deleteRow( required workbook,required numeric rowNum ){
		/* Deletes the data from a row. Does not physically delete the row. */
		var rowToDelete = rowNum - 1;
		if( rowToDelete GTE tools.getFirstRowNum( workbook ) AND rowToDelete LTE tools.getLastRowNum( workbook ) ) //If this is a valid row, remove it
			tools.getActiveSheet( workbook ).removeRow( tools.getActiveSheet( workbook ).getRow( JavaCast( "int",rowToDelete ) ) );
	}

	void function formatCell( required workbook,required struct format,required numeric row,required numeric column,any cellStyle ){
		var cell = tools.initializeCell( workbook,row,column );
		if( arguments.KeyExists( "cellStyle" ) )
			cell.setCellStyle( cellStyle );// reuse an existing style
		else
			cell.setCellStyle( formatting.buildCellStyle( workbook,format ) );
	}

	void function formatRow( required workbook,required struct format,required numeric rowNum ){
		var theRow = tools.getActiveSheet( workbook ).getRow( arguments.rowNum-1 );
		if( IsNull( theRow ) )
			return;
		var cellIterator = theRow.cellIterator();
		while( cellIterator.hasNext() ){
			formatCell( workbook,format,rowNum,cellIterator.next().getColumnIndex()+1 );
		}
	}

	function new( string sheetName="Sheet1",boolean xmlformat=false ){
		var workbook = tools.createWorkBook( sheetName.Left( 31 ) );
		tools.createSheet( workbook,sheetName,xmlformat );
		setActiveSheet( workbook,sheetName );
		return workbook;
	}

	function read(
		required string src
		,required string format
		,string columns
		,string columnnames=""
		,numeric headerrow
		,string rows
		,numeric sheet
		,string sheetname
		,boolean excludeHeaderRow=false
		,boolean readAllSheets=false
	){
		if( arguments.KeyExists( "query" ) )
			throw( type=exceptionType,message="Invalid argument 'query'.",details="Just use format='query' to return a query object" );
		if( arguments.KeyExists( "format" ) AND !ListFindNoCase( "query,csv,html,tab,pipe" ,format ) )
			throw( type=exceptionType,message="Invalid Format",detail="Supported formats are: QUERY, HTML, CSV, TAB and PIPE" );
		if( arguments.KeyExists( "sheetname" ) AND arguments.KeyExists( "sheet" ) )
			throw( type=exceptionType,message="Cannot Provide Both Sheet and sheetname Attributes",detail="Only one of either 'sheet' or 'sheetname' attributes may be provided." );
		var returnValue = 0;
		var exportUtil = 0;
		var outFile = "";
		/* create an exporter for the selected format */
		switch( format ){
			case "csv": case "tab": case "pipe":
				outFile = GetTempFile( ExpandPath( "." ),"cfpoi" );
				exportUtil = tools.loadPOI( "org.cfsearching.poi.WorkbookExportFactory" ).createCSVExport( src,outFile );
				exportUtil.setSeparator( exportUtil[ UCase( format ) ] );
				break;
			case "html":
				outFile = GetTempFile( ExpandPath( "." ),"cfpoi" );
				exportUtil = tools.loadPOI( "org.cfsearching.poi.WorkbookExportFactory" ).crecreateSimpleHTMLExportateCSVExport( src,outFile );
				break;
			case "query":
				exportUtil = tools.loadPOI( "org.cfsearching.poi.WorkbookExportFactory" ).createQueryExport( src,"q" );
				exportUtil.setColumnNames( JavaCast( "string",columnNames ) );
				break;
		}
		/* read a specific sheet */
		if( !readAllSheets ){
			if( arguments.KeyExists( "sheetname" ) )
				exportUtil.setSheetToRead( JavaCast( "string", sheetname ) );
			else if( arguments.KeyExists( "sheet" ) )
				exportUtil.setSheetToRead( JavaCast( "int",sheet-1 ) );
			else
				exportUtil.setSheetToRead( JavaCast( "int",0 ) );
		}
		/*  read a specific range of rows */
		if( arguments.KeyExists( "rows" ) )
			exportUtil.setRowsToProcess( JavaCast( "string",rows ) );
			/* read a specific range of columns */
		if( arguments.KeyExists( "columns" ) )
			exportUtil.setColumnsToProcess( JavaCast( "string",columns ) );
		/* identify header row */
		if( arguments.KeyExists( "headerrow" ) )
			exportUtil.setHeaderRow( JavaCast( "int",headerRow-1 ) );
		/* for ACF compatibility */
		if( arguments.KeyExists( "excludeHeaderRow" ) AND excludeHeaderRow )
			exportUtil.setExcludeHeaderRow( JavaCast( "boolean",true ) );
		try{
			exportUtil.process();
			if( format IS "query" )
				returnValue = exportUtil.getQuery();
			else
				returnValue = FileRead( outFile,"UTF-8" );
		}
		finally{
			if( FileExists( outFile ) )
				FileDelete( outfile );
		}
		return returnValue;
	}

	binary function readBinary( required workbook ){
		var baos = CreateObject( "Java","org.apache.commons.io.output.ByteArrayOutputStream" ).init();
		workbook.write( baos );
		baos.flush();
		return baos.toByteArray();
	}

	void function setActiveSheet( required workbook,string sheetName,numeric sheetIndex ){
		tools.validateSheetNameOrIndexWasProvided( argumentCollection=arguments );
		if( arguments.KeyExists( "sheetName" ) ){
			tools.validateSheetName( workbook,sheetName );
			arguments.sheetIndex = workbook.getSheetIndex( JavaCast( "string",arguments.sheetName ) ) + 1;
		}
		tools.validateSheetIndex( workbook,arguments.sheetIndex )
		workbook.setActiveSheet( JavaCast( "int",arguments.sheetIndex - 1 ) );
	}

	void function shiftRows( required workbook,required numeric startRow,numeric endRow=startRow,numeric offest=1 ){
		tools.getActiveSheet( workbook ).shiftRows(
			JavaCast( "int",arguments.startRow - 1 )
			,JavaCast( "int",arguments.endRow - 1 )
			,JavaCast( "int",arguments.offset )
		);
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
	function createSheet(){ notYetImplemented(); }
	function deleteColumn(){ notYetImplemented(); }
	function deleteColumns(){ notYetImplemented(); }
	function deleteRows(){ notYetImplemented(); }
	function formatCellRange(){ notYetImplemented(); }
	function formatColumn(){ notYetImplemented(); }
	function formatColumns(){ notYetImplemented(); }
	function formatRows(){ notYetImplemented(); }
	function getCellComment(){ notYetImplemented(); }
	function getCellFormula(){ notYetImplemented(); }
	function getCellValue(){ notYetImplemented(); }
	function info(){ notYetImplemented(); }
	function mergeCells(){ notYetImplemented(); }
	function removeSheet(){ notYetImplemented(); }
	function removeSheetNumber(){ notYetImplemented(); }
	function setActiveSheetNumber(){ notYetImplemented(); }
	function setCellComment(){ notYetImplemented(); }
	function setCellFormula(){ notYetImplemented(); }
	function setCellValue(){ notYetImplemented(); }
	function setColumnWidth(){ notYetImplemented(); }
	function setHeader(){ notYetImplemented(); }
	function setRowHeight(){ notYetImplemented(); }
	function shiftColumns(){ notYetImplemented(); }
	function write(){ notYetImplemented(); }

}