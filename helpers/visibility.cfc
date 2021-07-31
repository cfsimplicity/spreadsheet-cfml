component extends="base" accessors="true"{

	public void function doFillMergedCellsWithVisibleValue( required workbook, required sheet ){
		if( !getSheetHelper().sheetHasMergedRegions( arguments.sheet ) )
			return;
		for( var regionIndex = 0; regionIndex < arguments.sheet.getNumMergedRegions(); regionIndex++ ){
			var region = arguments.sheet.getMergedRegion( regionIndex );
			var regionStartRowNumber = ( region.getFirstRow() +1 );
			var regionEndRowNumber = ( region.getLastRow() +1 );
			var regionStartColumnNumber = ( region.getFirstColumn() +1 );
			var regionEndColumnNumber = ( region.getLastColumn() +1 );
			var visibleValue = library().getCellValue( arguments.workbook, regionStartRowNumber, regionStartColumnNumber );
			library().setCellRangeValue( arguments.workbook, visibleValue, regionStartRowNumber, regionEndRowNumber, regionStartColumnNumber, regionEndColumnNumber );
		}
	}

	public void function toggleColumnHidden( required workbook, required numeric columnNumber, required boolean state ){
		getSheetHelper().getActiveSheet( arguments.workbook ).setColumnHidden( JavaCast( "int", arguments.columnNumber-1 ), JavaCast( "boolean", arguments.state ) );
	}

	public void function toggleRowHidden( required workbook, required numeric rowNumber, required boolean state ){
		getRowHelper().getRowFromActiveSheet( arguments.workbook, arguments.rowNumber ).setZeroHeight( JavaCast( "boolean", arguments.state ) );
	}

}