component extends="base"{

	// underscore prefix because otherwise errors: "no matching function [autoSizeColumns]"
	void function _autoSizeColumns( required workbook, required numeric startColumnNumber, required numeric endColumnNumber ){
		for( var i = startColumnNumber; i <= endColumnNumber; i++ )
			library().autoSizeColumn( arguments.workbook, i );
	}

	numeric function columnCountFromRanges( required array ranges ){
		var result = 0;
		for( var thisRange in arguments.ranges ){
			for( var i = thisRange.startAt; i <= thisRange.endAt; i++ )
				result++;
		}
		return result;
	}

	void function shiftColumnsRightStartingAt( required numeric cellIndex, required row, required workbook ){
		var lastCellIndex = arguments.row.getLastCellNum()-1;
		for( var i = lastCellIndex; i >= arguments.cellIndex; i-- )
			getCellHelper().shiftCell( arguments.workbook, arguments.row, i, 1 );
	}

	any function toggleColumnHidden( required workbook, required numeric columnNumber, required boolean state ){
		getSheetHelper().getActiveSheet( arguments.workbook ).setColumnHidden( JavaCast( "int", arguments.columnNumber-1 ), JavaCast( "boolean", arguments.state ) );
		return this;
	}

	numeric function getNewColumnIndex( required sheet, required numeric rowIndex, required numeric startColumn ){
		if( arguments.startColumn > 0 )
			return ( arguments.startColumn -1 );
		var row = arguments.sheet.getRow( arguments.rowIndex );
		if( !IsNull( row ) && getRowHelper().rowHasCells( row ) )
			return getRowHelper().getNextEmptyCellIndexFromRow( row );// append the new column to the existing columns
		return 0;
	}

}