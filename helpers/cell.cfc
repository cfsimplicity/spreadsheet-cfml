component extends="base" accessors="true"{

	property name="cellUtil" getter="false" setter="false";

	any function getCellUtil(){
		if( IsNull( variables.cellUtil ) )
			variables.cellUtil = getClassHelper().loadClass( "org.apache.poi.ss.util.CellUtil" );
		return variables.cellUtil;
	}

	boolean function cellExists( required workbook, required numeric rowNumber, required numeric columnNumber ){
		var checkRow = getRowHelper().getRowFromActiveSheet( arguments.workbook, arguments.rowNumber );
		var columnIndex = ( arguments.columnNumber -1 );
		return !IsNull( checkRow ) && !IsNull( checkRow.getCell( JavaCast( "int", columnIndex ) ) );
	}

	boolean function cellIsOfType( required cell, required string type ){
		var cellType = arguments.cell.getCellType();
		return ObjectEquals( cellType, cellType[ arguments.type ] );
	}

	any function createCell( required row, numeric cellNum=arguments.row.getLastCellNum(), overwrite=true ){
		// get existing cell (if any)
		var cell = arguments.row.getCell( JavaCast( "int", arguments.cellNum ) );
		if( arguments.overwrite && !IsNull( cell ) )
			arguments.row.removeCell( cell );// forcibly remove the existing cell
		if( arguments.overwrite || IsNull( cell ) )
			cell = arguments.row.createCell( JavaCast( "int", arguments.cellNum ) );// create a brand new cell
		return cell;
	}

	any function getCellAt( required workbook, required numeric rowNumber, required numeric columnNumber ){
		if( !cellExists( argumentCollection=arguments ) )
			Throw( type=library().getExceptionType(), message="Invalid cell", detail="The requested cell [#arguments.rowNumber#,#arguments.columnNumber#] does not exist in the active sheet" );
		var columnIndex = ( arguments.columnNumber -1 );
		return getRowHelper().getRowFromActiveSheet( arguments.workbook, arguments.rowNumber ).getCell( JavaCast( "int", columnIndex ) );
	}

	any function getCellFormulaValue( required workbook, required cell ){
		var formulaEvaluator = arguments.workbook.getCreationHelper().createFormulaEvaluator();
		try{
			return getFormatHelper().getDataFormatter().formatCellValue( arguments.cell, formulaEvaluator );
		}
		catch( any exception ){
			Throw( type=library().getExceptionType(), message="Failed to run formula", detail="There is a problem with the formula in sheet #arguments.cell.getSheet().getSheetName()# row #( arguments.cell.getRowIndex() +1 )# column #( arguments.cell.getColumnIndex() +1 )#");
		}
	}

	any function getCellRangeAddressFromColumnAndRowIndices( required struct indices ){
		//index = 0 based
		return getClassHelper().loadClass( "org.apache.poi.ss.util.CellRangeAddress" ).init(
			JavaCast( "int", arguments.indices.startRow )
			,JavaCast( "int", arguments.indices.endRow )
			,JavaCast( "int", arguments.indices.startColumn )
			,JavaCast( "int", arguments.indices.endColumn )
		);
	}

	any function getCellRangeAddressFromReference( required string rangeReference ){
		/*
		rangeReference = usually a standard area ref (e.g. "B1:D8"). May be a single cell ref (e.g. "B5") in which case the result is a 1 x 1 cell range. May also be a whole row range (e.g. "3:5"), or a whole column range (e.g. "C:F")
		*/
		return getClassHelper().loadClass( "org.apache.poi.ss.util.CellRangeAddress" ).valueOf( JavaCast( "String", arguments.rangeReference ) );
	}

	any function getCellValueAsType( required workbook, required cell ){
		/*
		Get the value of the cell based on the data type. The thing to worry about here is cell formulas and cell dates. Formulas can be strange and dates are stored as numeric types. Here I will just grab dates as floats and formulas I will try to grab as numeric values.
		*/
		if( cellIsOfType( arguments.cell, "NUMERIC" ) )
			return getCellNumericOrDateValue( arguments.cell );
		if( cellIsOfType( arguments.cell, "FORMULA" ) )
			return getCellFormulaValue( arguments.workbook, arguments.cell );
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

	any function initializeCell( required workbook, required numeric rowNumber, required numeric columnNumber ){
		//Automatically creates the cell if it does not exist, instead of throwing an error
		var rowIndex = JavaCast( "int", ( arguments.rowNumber -1 ) );
		var columnIndex = JavaCast( "int", ( arguments.columnNumber -1 ) );
		var rowObject = getCellUtil().getRow( rowIndex, getSheetHelper().getActiveSheet( arguments.workbook ) );
		var cellObject = getCellUtil().getCell( rowObject, columnIndex );
		return cellObject;
	}

	any function setCellValueAsType( required workbook, required cell, required value, string type ){
		var validCellTypes = [ "string", "numeric", "date", "time", "boolean", "blank" ];
		if( !arguments.KeyExists( "type" ) ) //autodetect type
			arguments.type = getDataTypeHelper().detectValueDataType( arguments.value );
		else if( !validCellTypes.FindNoCase( arguments.type ) )
			Throw( type=library().getExceptionType(), message="Invalid data type: '#arguments.type#'", detail="The data type must be one of the following: #validCellTypes.ToList( ', ' )#." );
		/* Note: To properly apply date/number formatting:
			- cell type must be CELL_TYPE_NUMERIC (NB: POI5+ can't set cell types explicitly anymore: https://bz.apache.org/bugzilla/show_bug.cgi?id=63118 )
			- cell value must be applied as a java.util.Date or java.lang.Double (NOT as a string)
			- cell style must have a dataFormat (datetime values only)
 		*/
		switch( arguments.type ){
			case "numeric":
				return setNumericValue( argumentCollection=arguments );
			case "date": case "time":
				return setDateOrTimeValue( argumentCollection=arguments );
			case "boolean":
				return setBooleanValue( argumentCollection=arguments );
			case "blank":
				return setEmptyValue( argumentCollection=arguments );
		}
		return setStringValue( argumentCollection=arguments );
	}

	any function shiftCell( required workbook, required row, required numeric cellIndex, required numeric offset ){
		var originalCell = arguments.row.getCell( JavaCast( "int", arguments.cellIndex ) );
		if( IsNull( originalCell ) )
			return this;
		var cell = createCell( arguments.row, arguments.cellIndex + arguments.offset );
		setCellValueAsType( arguments.workbook, cell, getCellValueAsType( arguments.workbook, originalCell ) );
		cell.setCellStyle( originalCell.getCellStyle() );
		cell.setCellComment( originalCell.getCellComment() );
		cell.setHyperlink( originalCell.getHyperLink() );
		arguments.row.removeCell( originalCell );
		return this;
	}

	/* PRIVATE */
	private boolean function isCellDateFormated( required any cell ){
		return getDateHelper().getDateUtil().isCellDateFormatted( arguments.cell );
	}

	private any function getCellNumericOrDateValue( required any cell ){
		// Get numeric cell data. This could be a standard number, could also be a date value.
		if( !isCellDateFormated( arguments.cell ) )
			return arguments.cell.getNumericCellValue();
		getDateHelper().matchPoiTimeZoneToEngine();
		var cellValue = arguments.cell.getDateCellValue();
		if( getDateHelper().isTimeOnlyValue( cellValue ) )
			return getFormatHelper().getDataFormatter().formatCellValue( arguments.cell );//return as a time formatted string to avoid default epoch date 1899-12-31
		return cellValue;
	}

	private any function setNumericValue( required any cell, required any value ){
		arguments.cell.setCellValue( JavaCast( "double", Val( arguments.value ) ) );
		return this;
	}

	private any function setDateOrTimeValue( required workbook, required cell, required value, string type ){
		getDateHelper().matchPoiTimeZoneToEngine();
		//handle empty strings which can't be treated as dates
		if( Trim( arguments.value ).IsEmpty() ){
			arguments.cell.setBlank(); //no need to set the value: it will be blank
			return this;
		}
		var dateTimeValue = ParseDateTime( arguments.value );
		if( arguments.type == "time" )
			var cellFormat = library().getDateFormats().TIME; //don't include the epoch date in the display
		else
			var cellFormat = getDateHelper().getDefaultDateMaskFor( dateTimeValue );// check if DATE, TIME or TIMESTAMP
		var dataFormat = arguments.workbook.getCreationHelper().createDataFormat();
		//Use setCellStyleProperty() which will try to re-use an existing style rather than create a new one for every cell which may breach the 4009 styles per wookbook limit
		getCellUtil().setCellStyleProperty( arguments.cell, getCellUtil().DATA_FORMAT, dataFormat.getFormat( JavaCast( "string", cellFormat ) ) );
		/*  Excel uses a different epoch than CF (1900-01-01 versus 1899-12-30). "Time" only values will not display properly without special handling */
		if( arguments.type == "time" || getDateHelper().isTimeOnlyValue( dateTimeValue ) ){
			dateTimeValue = dateTimeValue.Add( "d", 2 );//shift the epoch forward to match Excel's
			var javaDate = dateTimeValue.from( dateTimeValue.toInstant() );// dateUtil needs a java date
			dateTimeValue = ( getDateHelper().getDateUtil().getExcelDate( javaDate ) -1 );//Convert to Excel's double value for dates, minus the 1 complete day to leave the day fraction (= time value)
		}
		arguments.cell.setCellValue( dateTimeValue );
		return this;
	}

	private any function setBooleanValue( required any cell, required any value ){
		//handle empty strings/nulls which can't be treated as booleans
		if( Trim( arguments.value ).IsEmpty() ){
			arguments.cell.setBlank(); //no need to set the value: it will be blank
			return this;
		}
		arguments.cell.setCellValue( JavaCast( "boolean", arguments.value ) );
		return this;
	}

	private any function setEmptyValue( required any cell ){
		arguments.cell.setBlank(); //no need to set the value: it will be blank
		return this;
	}

	private any function setStringValue( required any cell, required any value ){
		arguments.cell.setCellValue( JavaCast( "string", arguments.value ) );
		return this;
	}

}