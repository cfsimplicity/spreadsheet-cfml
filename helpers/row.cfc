component extends="base" accessors="true"{

	public any function addRowToSheetData(
		required workbook
		,required struct sheet
		,required numeric rowIndex
		,boolean includeRichTextFormatting=false
	){
		if( ( arguments.rowIndex == arguments.sheet.headerRowIndex ) && !arguments.sheet.includeHeaderRow )
			return this;
		var rowData = [];
		var row = arguments.sheet.object.getRow( JavaCast( "int", arguments.rowIndex ) );
		if( IsNull( row ) ){
			if( arguments.sheet.includeBlankRows )
				arguments.sheet.data.Append( rowData );
			return this;
		}
		if( rowIsEmpty( row ) && !arguments.sheet.includeBlankRows )
			return this;
		rowData = getRowData( arguments.workbook, row, arguments.sheet.columnRanges, arguments.includeRichTextFormatting );
		arguments.sheet.data.Append( rowData );
		if( !arguments.sheet.columnRanges.Len() ){
			var rowColumnCount = row.getLastCellNum();
			arguments.sheet.totalColumnCount = Max( arguments.sheet.totalColumnCount, rowColumnCount );
		}
		return this;
	}

	public any function createRow( required workbook, numeric rowNum=getNextEmptyRowNumber( arguments.workbook ), boolean overwrite=true ){
		// get existing row (if any)
		var sheet = getSheetHelper().getActiveSheet( arguments.workbook );
		var row = sheet.getRow( JavaCast( "int", arguments.rowNum ) );
		if( arguments.overwrite && !IsNull( row ) )
			sheet.removeRow( row ); // forcibly remove existing row and all cells
		if( arguments.overwrite || IsNull( sheet.getRow( JavaCast( "int", arguments.rowNum ) ) ) ){
			try{
				row = sheet.createRow( JavaCast( "int", arguments.rowNum ) );
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

	public numeric function getFirstRowIndex( required workbook ){
		var sheet = getSheetHelper().getActiveSheet( arguments.workbook );
		var firstRow = sheet.getFirstRowNum();
		if( ( firstRow == 0 ) && ( sheet.getPhysicalNumberOfRows() == 0 ) )
			return -1;
		return firstRow;
	}

	public numeric function getLastRowIndex( required workbook, sheet=getSheetHelper().getActiveSheet( arguments.workbook ) ){
		var lastRow = arguments.sheet.getLastRowNum();
		if( ( lastRow == 0 ) && ( arguments.sheet.getPhysicalNumberOfRows() == 0 ) )
			return -1; //The sheet is empty. Return -1 instead of 0
		return lastRow;
	}

	public numeric function getNextEmptyCellIndexFromRow( required row ){
		return arguments.row.getLastCellNum(); //getLastCellNum() = the last cell index +1
	}

	public numeric function getNextEmptyRowNumber( workbook ){
		return ( getLastRowIndex( arguments.workbook ) +1 );
	}

	public array function getRowData( required workbook, required row, array columnRanges=[], boolean includeRichTextFormatting=false ){
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

	public any function getRowFromActiveSheet( required workbook, required numeric rowNumber ){
		var rowIndex = ( arguments.rowNumber-1 );
		return getSheetHelper().getActiveSheet( arguments.workbook ).getRow( JavaCast( "int", rowIndex ) );
	}

	public array function parseListDataToArray( required string line, required string delimiter, boolean handleEmbeddedCommas=true ){
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

	public boolean function rowHasCells( required row ){
		return ( arguments.row.getLastCellNum() > 0 );
	}

	public any function shiftOrDeleteRow(
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

	/* Private */

	private boolean function rowIsEmpty( required row ){
		for( var i = arguments.row.getFirstCellNum(); i < arguments.row.getLastCellNum(); i++ ){
	    var cell = arguments.row.getCell( i );
	    if( !IsNull( cell ) && !getCellHelper().cellIsOfType( cell, "BLANK" ) )
	    	return false;
	  }
	  return true;
	}

}