component extends="base" accessors="true"{

	any function addRowToSheetData(
		required workbook
		,required struct sheet
		,required numeric rowIndex
		,boolean includeRichTextFormatting=false
	){
		if( ( arguments.rowIndex == arguments.sheet.headerRowIndex ) && !arguments.sheet.includeHeaderRow ){
			var row = arguments.sheet.object.getRow( JavaCast( "int", arguments.rowIndex ) );
			setSheetColumnCountFromRow( row, arguments.sheet );
			return this;
		}
		var rowData = [];
		var row = arguments.sheet.object.getRow( JavaCast( "int", arguments.rowIndex ) );
		if( IsNull( row ) ){
			if( arguments.sheet.includeBlankRows )
				arguments.sheet.data.Append( rowData );
			return this;
		}
		if( rowIsEmpty( row ) && !arguments.sheet.includeBlankRows )
			return this;
		if( rowIsHidden( row ) && !arguments.sheet.includeHiddenRows )
			return this;
		rowData = getRowData( arguments.workbook, row, arguments.sheet.columnRanges, arguments.includeRichTextFormatting );
		arguments.sheet.data.Append( rowData );
		setSheetColumnCountFromRow( row, arguments.sheet );
		return this;
	}

	any function createRow( required workbook, numeric rowIndex, boolean overwrite=true ){
		// get existing row (if any)
		var sheet = getSheetHelper().getActiveSheet( arguments.workbook );
		if( !arguments.KeyExists( "rowIndex" ) )
			arguments.rowIndex = getSheetHelper().getNextEmptyRowIndex( sheet );
		var row = sheet.getRow( JavaCast( "int", arguments.rowIndex ) );
		if( arguments.overwrite && !IsNull( row ) )
			sheet.removeRow( row ); // forcibly remove existing row and all cells
		if( arguments.overwrite || IsNull( sheet.getRow( JavaCast( "int", arguments.rowIndex ) ) ) ){
			try{
				row = sheet.createRow( JavaCast( "int", arguments.rowIndex ) );
			}
			catch( java.lang.IllegalArgumentException exception ){
				if( exception.message.FindNoCase( "Invalid row number (65536)" ) )
					Throw( type=library().getExceptionType(), message="Too many rows", detail="Binary spreadsheets are limited to 65535 rows. Consider using an XML format spreadsheet instead." );
				else
					rethrow;
			}
		}
		return row;
	}

	numeric function getNextEmptyCellIndexFromRow( required row ){
		//NB: getLastCellNum() == the last cell index PLUS ONE or -1 if no cells
		if( arguments.row.getLastCellNum() == -1 )
			return 0;
		return arguments.row.getLastCellNum();
	}

	array function getRowData( required workbook, required row, array columnRanges=[], boolean includeRichTextFormatting=false ){
		var result = [];
		if( !arguments.columnRanges.Len() ){
			var columnRange = {
				startAt: 1
				,endAt: arguments.row.getLastCellNum()
			};
			arguments.columnRanges = [ columnRange ];
		}
		for( var thisRange in arguments.columnRanges ){
			for( var i = thisRange.startAt; i <= thisRange.endAt; i++ ){
				var colIndex = ( i-1 );
				var cell = arguments.row.GetCell( JavaCast( "int", colIndex ) );
				if( IsNull( cell ) ){
					result.Append( "" );
					continue;
				}
				var cellValue = getCellHelper().getCellValueAsType( arguments.workbook, cell );
				if( arguments.includeRichTextFormatting && getCellHelper().cellIsOfType( cell, "STRING" ) )
					cellValue = getFormatHelper().richStringCellValueToHtml( arguments.workbook, cell,cellValue );
				result.Append( cellValue );
			}
		}
		return result;
	}

	any function getRowFromActiveSheet( required workbook, required numeric rowNumber ){
		var rowIndex = ( arguments.rowNumber-1 );
		return getSheetHelper().getActiveSheet( arguments.workbook ).getRow( JavaCast( "int", rowIndex ) );
	}

	array function parseListDataToArray( required string line, required string delimiter, boolean handleEmbeddedCommas=true ){
		var elements = ListToArray( arguments.line, arguments.delimiter );
		var potentialQuotes = 0;
		arguments.line = ToString( arguments.line );
		if( ( arguments.delimiter == "," ) && arguments.handleEmbeddedCommas )
			potentialQuotes = arguments.line.ReplaceAll( "[^']", "" ).length();
		if( potentialQuotes <= 1 )
			return elements;
		//For ACF compatibility, find any values enclosed in single quotes and treat them as a single element.
		var currentValue = 0;
		var nextValue = "";
		var isEmbeddedValue = false;
		var values = [];
		var buffer = getStringHelper().newJavaStringBuilder();
		var maxElements = ArrayLen( elements );
		for( var i = 1; i <= maxElements; i++ ) {
		  currentValue = Trim( elements[ i ] );
		  nextValue = i < maxElements ? elements[ i + 1 ] : "";
		  var isComplete = false;
		  var hasLeadingQuote = ( currentValue.Left( 1 ) == "'" );
		  var hasTrailingQuote = ( currentValue.Right( 1 ) == "'" );
		  var isFinalElement = ( i == maxElements );
		  if( hasLeadingQuote )
		  	isEmbeddedValue = true;
		  if( isEmbeddedValue && hasTrailingQuote )
		  	isComplete = true;
		  /* We are finished with this value if:
			  * no quotes were found OR
			  * it is the final value OR
			  * the next value is embedded in quotes
		  */
		  if( !isEmbeddedValue || isFinalElement || ( nextValue.Left( 1 ) == "'" ) )
		  	isComplete = true;
		  if( isEmbeddedValue || isComplete ){
			  // if this a partial value, append the delimiter
			  if( isEmbeddedValue && buffer.length() > 0 )
			  	buffer.Append( "," );
			  buffer.Append( elements[ i ] );
		  }
		  if( isComplete ){
			  var finalValue = buffer.toString();
			  var startAt = finalValue.indexOf( "'" );
			  var endAt = finalValue.lastIndexOf( "'" );
			  if( isEmbeddedValue && startAt >= 0 && endAt > startAt )
			  	finalValue = finalValue.substring( ( startAt +1 ), endAt );
			  values.Append( finalValue );
			  buffer.setLength( 0 );
			  isEmbeddedValue = false;
		  }
	  }
	  return values;
	}

	any function populateFromQueryRow(
		required workbook
		,required newRow
		,required rowData
		,required array queryColumns
		,required numeric firstCellIndex
		,required boolean ignoreQueryColumnDataTypes
	){
		var cellIndex = arguments.firstCellIndex;
		var overrideDataTypes = arguments.KeyExists( "datatypes" );
		for( var queryColumn in queryColumns ){
 			var cell = getCellHelper().createCell( arguments.newRow, cellIndex, false );
			var cellValue = rowData[ queryColumn.name ];
			if( arguments.ignoreQueryColumnDataTypes ){
				if( overrideDataTypes )
   				getDataTypeHelper().setCellDataTypeWithOverride( arguments.workbook, cell, cellValue, cellIndex, arguments.datatypes );
   			else
					getCellHelper().setCellValueAsType( arguments.workbook, cell, cellValue );
				cellIndex++;
				continue;
			}
			var cellValueType = getDataTypeHelper().getCellValueTypeFromQueryColumnType( queryColumn.cellDataType, cellValue );
			if( overrideDataTypes )
 				getDataTypeHelper().setCellDataTypeWithOverride( arguments.workbook, cell, cellValue, cellIndex, arguments.datatypes, cellValueType );
 			else
				getCellHelper().setCellValueAsType( arguments.workbook, cell, cellValue, cellValueType );
			cellIndex++;
 		}
 		return this;
	}

	numeric function populateFromArray(
		required workbook
		,required newRow
		,required array rowData
		,required numeric firstCellIndex
		,required numeric currentMaxColumnCount
	){
		var cellIndex = arguments.firstCellIndex;
		var columnCount = arguments.currentMaxColumnCount;
		var overrideDataTypes = arguments.KeyExists( "datatypes" );
		cfloop( array=rowData, item="local.cellValue", index="local.thisColumnNumber" ){
 			var cell = getCellHelper().createCell( arguments.newRow, cellIndex );
 			if( overrideDataTypes )
 				getDataTypeHelper().setCellDataTypeWithOverride( arguments.workbook, cell, cellValue, cellIndex, arguments.datatypes );
 			else
				getCellHelper().setCellValueAsType( arguments.workbook, cell, cellValue );
			columnCount = Max( columnCount, thisColumnNumber );
			cellIndex++;
		}
		return columnCount;
	}

	boolean function rowHasCells( required row ){
		return ( arguments.row.getLastCellNum() > 0 );
	}

	any function shiftOrDeleteRow(
		required workbook
		,required row
		,required lastRow
		,required boolean insert
	){
		if( arguments.insert ){
			library().shiftRows( arguments.workbook, arguments.row, arguments.lastRow, 1 );//shift the existing rows down (by one row)
			return this;
		}
		library().deleteRow( arguments.workbook, arguments.row );//otherwise, clear the entire row
		return this;
	}

	boolean function isRowHidden( required workbook, required numeric row ){
		return rowIsHidden( getRowFromActiveSheet( arguments.workbook, arguments.row ) );
	}

	any function toggleRowHidden( required workbook, required numeric rowNumber, required boolean state ){
		getRowFromActiveSheet( arguments.workbook, arguments.rowNumber ).setZeroHeight( JavaCast( "boolean", arguments.state ) );
		return this;
	}

	/* Private */

	private boolean function rowIsEmpty( required row ){
		for( var i = arguments.row.getFirstCellNum(); i < arguments.row.getLastCellNum(); i++ ){
	    var cell = arguments.row.getCell( i );
	    if( !IsNull( cell ) && !getCellHelper().cellIsOfType( cell, "BLANK" ) )
	    	return false;
	  }
	  return true;
	}

	private boolean function rowIsHidden( required row ){
		return arguments.row.getZeroHeight() || arguments.row.getHeight() == 0;
	}

	private void function setSheetColumnCountFromRow( required any row, required struct sheet ){
		if( arguments.sheet.columnRanges.Len() )//don't change the column count if specific columns have been specified
			return;
		arguments.sheet.totalColumnCount = Max( arguments.sheet.totalColumnCount, arguments.row.getLastCellNum() );
	}

}