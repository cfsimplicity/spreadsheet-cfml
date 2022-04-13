component extends="base" accessors="true"{

	string function createOrValidateSheetName( required workbook ){
		if( !arguments.KeyExists( "sheetName" ) )
			return generateUniqueSheetName( arguments.workbook );
		validateSheetName( arguments.sheetName );
		return arguments.sheetName;
	}

	any function deleteSheetAtIndex( required workbook, required numeric sheetIndex ){
		arguments.workbook.removeSheetAt( JavaCast( "int", arguments.sheetIndex ) );
		return this;
	}

	any function getActiveSheet( required workbook ){
		return arguments.workbook.getSheetAt( JavaCast( "int", arguments.workbook.getActiveSheetIndex() ) );
	}

	any function getActiveSheetFooter( required workbook ){
		return getActiveSheet( arguments.workbook ).getFooter();
	}

	any function getActiveSheetHeader( required workbook ){
		return getActiveSheet( arguments.workbook ).getHeader();
	}

	any function getActiveSheetName( required workbook ){
		return getActiveSheet( arguments.workbook ).getSheetName();
	}

	any function setActiveSheetNameOrNumber( required workbook, required sheetNameOrNumber ){
		if( IsValid( "integer", arguments.sheetNameOrNumber ) && IsNumeric( arguments.sheetNameOrNumber ) ){
			var sheetNumber = arguments.sheetNameOrNumber;
			library().setActiveSheetNumber( arguments.workbook, sheetNumber );
			return this;
		}
		var sheetName = arguments.sheetNameOrNumber;
		library().setActiveSheet( arguments.workbook, sheetName );
		return this;
	}

	array function getAllSheetFormulas( required workbook ){
		var rowIterator = getActiveSheet( arguments.workbook ).rowIterator();
		var formulas = [];
		while( rowIterator.hasNext() ){
			var cellIterator = rowIterator.next().cellIterator();
			while( cellIterator.hasNext() ){
				var cell = cellIterator.next();
				var cellFormula = getCellFormula( cell );
				if ( cellFormula.Len() ) {
					formulas.Append( {
						row: ( cell.getRowIndex() + 1 )
						,column: ( cell.getColumnIndex() + 1 )
						,formula: cellFormula
					} );
				}
			}
		}
		return formulas;
	}

	numeric function getFirstRowIndex( required sheet ){
		return arguments.sheet.getFirstRowNum(); //-1 if no rows exist
	}

	numeric function getLastRowIndex( required sheet ){
		return arguments.sheet.getLastRowNum(); //-1 if no rows exist
	}

	numeric function getNextEmptyRowIndex( required sheet ){
		return ( getLastRowIndex( arguments.sheet ) +1 );
	}

	any function getSheetByName( required workbook, required string sheetName ){
		validateSheetExistsWithName( arguments.workbook, arguments.sheetName );
		return arguments.workbook.getSheet( JavaCast( "string", arguments.sheetName ) );
	}

	any function getSheetByNumber( required workbook, required numeric sheetNumber ){
		validateSheetNumber( arguments.workbook, arguments.sheetNumber );
		var sheetIndex = ( arguments.sheetNumber -1 );
		return arguments.workbook.getSheetAt( sheetIndex );
	}

	any function getSpecifiedOrActiveSheet( required workbook, string sheetName, numeric sheetNumber ){
		throwErrorIFSheetNameAndNumberArgumentsBothPassed( argumentCollection=arguments );
		if( !sheetNameArgumentWasProvided( argumentCollection=arguments ) && !sheetNumberArgumentWasProvided( argumentCollection=arguments ) )
			return getActiveSheet( arguments.workbook );
		if( sheetNameArgumentWasProvided( argumentCollection=arguments ) )
			return getSheetByName( arguments.workbook, arguments.sheetName );
		return getSheetByNumber( arguments.workbook, arguments.sheetNumber );
	}

	struct function info( required workbook, numeric sheetNumber ){
		if( !arguments.KeyExists( "sheetNumber" ) )
			arguments.sheetNumber = ( arguments.workbook.getActiveSheetIndex() +1 );
		var sheet = getSheetByNumber( argumentCollection=arguments );
		var isXlsx = library().isXmlFormat( arguments.workbook );
		return {
			displaysAutomaticPageBreaks: sheet.getAutobreaks()
			,displaysFormulas: sheet.isDisplayFormulas()
			,displaysGridlines: sheet.isDisplayGridlines()
			,displaysRowAndColumnHeadings: sheet.isDisplayRowColHeadings()
			,displaysZeros: sheet.isDisplayZeros()
			,hasComments: hasComments( sheet, isXlsx )
			,hasMergedRegions: BooleanFormat( sheet.getNumMergedRegions() )
			,isCurrentActiveSheet: isActive( sheet, isXlsx )
			,isHidden: !isVisible( argumentCollection=arguments )
			,isRightToLeft: sheet.isRightToLeft()
			,name: sheet.getSheetName()
			,numberOfMergedRegions: sheet.getNumMergedRegions()
			,printsFitToPage: sheet.getFitToPage()
			,printsGridlines: sheet.isPrintGridlines()
			,printsHorizontallyCentered: sheet.getHorizontallyCenter()
			,printsRowAndColumnHeadings: sheet.isPrintRowAndColumnHeadings()
			,printsVerticallyCentered: sheet.getVerticallyCenter()
			,recalculateFormulasOnNextOpen: sheet.getForceFormulaRecalculation()
			,visibility: getVisibility( argumentCollection=arguments )
		};
	}

	any function moveSheet( required workbook, required string sheetName, required string moveToIndex ){
		arguments.workbook.setSheetOrder( JavaCast( "String", arguments.sheetName ), JavaCast( "int", arguments.moveToIndex ) );
		return this;
	}

	void function setVisibility( required workbook, required numeric sheetNumber, required string visibility ){
		/* POI Docs: "Please note that the sheet currently set as active sheet (sheet 0 in a newly created workbook or the one set via setActiveSheet()) cannot be hidden." */
		validateSheetNumber( arguments.workbook, arguments.sheetNumber );
		var validStates = [ "HIDDEN", "VERY_HIDDEN", "VISIBLE" ];
		if( !validStates.Find( arguments.visibility ) )
			Throw( type=this.getExceptionType(), message="Invalid visibility parameter: '#arguments.visibility#'", detail="The visibility must be one of the following: #validStates.ToList( ', ' )#." );
		var visibilityEnum = getClassHelper().loadClass( "org.apache.poi.ss.usermodel.SheetVisibility" )[ JavaCast( "string", arguments.visibility ) ];
		var sheetIndex = ( arguments.sheetNumber -1 );
		arguments.workbook.setSheetVisibility( sheetIndex, visibilityEnum );
	}

	boolean function isVisible( required workbook, required numeric sheetNumber ){
		return ( getVisibility( argumentCollection=arguments ) == "VISIBLE" );
	}

	boolean function sheetExists( required workbook, string sheetName, numeric sheetNumber ){
		validateSheetNameOrNumberWasProvided( argumentCollection=arguments );
		if( arguments.KeyExists( "sheetName" ) )
			arguments.sheetNumber = ( getSheetIndexFromName( arguments.workbook, arguments.sheetName ) +1 );
			//the position is valid if it's an integer between 1 and the total number of sheets in the workbook
		if( arguments.sheetNumber && ( arguments.sheetNumber == Round( arguments.sheetNumber ) ) && ( arguments.sheetNumber <= arguments.workbook.getNumberOfSheets() ) )
			return true;
		return false;
	}

	boolean function hasMergedRegions( required sheet ){
		return ( arguments.sheet.getNumMergedRegions() > 0 );
	}

	query function sheetToQuery(
		required workbook
		,string sheetName
		,numeric sheetNumber
		,numeric headerRow
		,boolean includeHeaderRow=false
		,boolean includeBlankRows=false
		,boolean includeHiddenColumns=false
		,boolean fillMergedCellsWithVisibleValue=false
		,boolean includeRichTextFormatting=false
		,string rows //range
		,string columns //range
		,any columnNames //list or array
		,any queryColumnTypes="" //'auto', single default type e.g. 'VARCHAR', or list of types, or struct of column names/types mapping. Empty means no types are specified.
		,boolean makeColumnNamesSafe=false
	){
		var sheet = {
			includeHeaderRow: arguments.includeHeaderRow
			,hasHeaderRow: ( arguments.KeyExists( "headerRow" ) && Val( arguments.headerRow ) )
			,includeBlankRows: arguments.includeBlankRows
			,columnNames: []
			,columnRanges: []
			,totalColumnCount: 0
			,data: []
		};
		if( arguments.KeyExists( "columnNames" ) && arguments.columnNames.Len() )
			sheet.columnNames = IsArray( arguments.columnNames )? arguments.columnNames: arguments.columnNames.ListToArray();
		sheet.headerRowIndex = sheet.hasHeaderRow? ( arguments.headerRow -1 ): -1;
		if( arguments.KeyExists( "columns" ) ){
			sheet.columnRanges = getRangeHelper().extractRanges( arguments.columns, arguments.workbook, "column" );
			sheet.totalColumnCount = getColumnHelper().columnCountFromRanges( sheet.columnRanges );
		}
		if( arguments.KeyExists( "sheetName" ) ){
			validateSheetExistsWithName( arguments.workbook, arguments.sheetName );
			arguments.sheetNumber = ( getSheetIndexFromName( arguments.workbook, arguments.sheetName ) +1 );
		}
		else if( !arguments.KeyExists( "sheetNumber" ) )
			arguments.sheetNumber = getFirstVisibleSheetNumber( arguments.workbook );
		if( arguments.sheetNumber == 0 )
			return QueryNew( "" );//no visible sheets
		sheet.object = getSheetByNumber( arguments.workbook, arguments.sheetNumber );
		var sheetHasRows = !sheetIsEmpty( sheet.object );
		if( sheetHasRows ){
			if( arguments.fillMergedCellsWithVisibleValue )
				doFillMergedCellsWithVisibleValue( arguments.workbook, sheet.object );
			if( arguments.KeyExists( "rows" ) ){
				var allRanges = getRangeHelper().extractRanges( arguments.rows, arguments.workbook );
				for( var thisRange in allRanges ){
					for( var rowNumber = thisRange.startAt; rowNumber <= thisRange.endAt; rowNumber++ ){
						var rowIndex = ( rowNumber -1 );
						getRowHelper().addRowToSheetData( arguments.workbook, sheet, rowIndex, arguments.includeRichTextFormatting );
					}
				}
			}
			else{
				var lastRowIndex = sheet.object.getLastRowNum();// zero based
				for( var rowIndex = 0; rowIndex <= lastRowIndex; rowIndex++ )
					getRowHelper().addRowToSheetData( arguments.workbook, sheet, rowIndex, arguments.includeRichTextFormatting );
			}
		}
		generateQueryColumnNames( arguments.workbook, sheet );
		arguments.queryColumnTypes = getQueryHelper().parseQueryColumnTypesArgument( arguments.queryColumnTypes, sheet.columnNames, sheet.totalColumnCount, sheet.data );
		var result = getQueryHelper()._QueryNew( sheet.columnNames, arguments.queryColumnTypes, sheet.data, arguments.makeColumnNamesSafe );
		if( !arguments.includeHiddenColumns && sheetHasRows ){
			result = getQueryHelper().deleteHiddenColumnsFromQuery( sheet, result );
			if( sheet.totalColumnCount == 0 )
				return QueryNew( "" );// all columns were hidden: return a blank query.
		}
		return result;
	}

	any function validateSheetExistsWithName( required workbook, required string sheetName ){
		if( !sheetExists( workbook=arguments.workbook, sheetName=arguments.sheetName ) )
			Throw( type=library().getExceptionType(), message="Invalid sheet name [#arguments.sheetName#]", detail="The specified sheet was not found in the current workbook." );
		return this;
	}

	any function validateSheetNumber( required workbook, required numeric sheetNumber ){
		if( !sheetExists( workbook=arguments.workbook, sheetNumber=arguments.sheetNumber ) ){
			var sheetCount = arguments.workbook.getNumberOfSheets();
			Throw( type=library().getExceptionType(), message="Invalid sheet number [#arguments.sheetNumber#]", detail="The sheetNumber must a whole number between 1 and the total number of sheets in the workbook [#sheetCount#]" );
		}
		return this;
	}

	any function validateSheetName( required string sheetName ){
		var characterCount = Len( arguments.sheetName );
		if( characterCount > 31 )
			Throw( type=library().getExceptionType(), message="Invalid sheet name", detail="The sheetname contains too many characters [#characterCount#]. The maximum is 31." );
		var poiTool = getClassHelper().loadClass( "org.apache.poi.ss.util.WorkbookUtil" );
		try{
			poiTool.validateSheetName( JavaCast( "String", arguments.sheetName ) );
		}
		catch( "java.lang.IllegalArgumentException" exception ){
			Throw( type=library().getExceptionType(), message="Invalid characters in sheet name", detail=exception.message );
		}
		catch( "java.lang.reflect.InvocationTargetException" exception ){
			//ACF
			Throw( type=library().getExceptionType(), message="Invalid characters in sheet name", detail=exception.message );
		}
		return this;
	}

	void function validateSheetNameOrNumberWasProvided(){
		throwErrorIFSheetNameAndNumberArgumentsBothMissing( argumentCollection=arguments );
		throwErrorIFSheetNameAndNumberArgumentsBothPassed( argumentCollection=arguments );
	}

	/* Private */


	private string function getCellFormula( required cell ) {
		if ( getCellHelper().cellIsOfType( cell, "FORMULA" ) ) {
			return cell.getCellFormula();
		}
		return "";
	}

	private any function generateQueryColumnNames( required workbook, required struct sheet ){
		if( arguments.sheet.columnNames.Len() ){
			forceQueryColumnsToMatchSpecifiedColumns( arguments.sheet );
			return this; // already generated
		}
		if( sheetIsEmpty( arguments.sheet.object ) )
			return this;
		if( arguments.sheet.hasHeaderRow ){
			// use specified header row values as column names
			var headerRowObject = arguments.sheet.object.getRow( JavaCast( "int", arguments.sheet.headerRowIndex ) );
			var headerRowData = getRowHelper().getRowData( arguments.workbook, headerRowObject, arguments.sheet.columnRanges );
			// adds default column names if header row column count is less than total data column count
			cfloop( from=1, to=arguments.sheet.totalColumnCount, index="local.i" ){
				arguments.sheet.columnNames.Append( getQueryColumnNameFromSpecifiedNames( headerRowData, i ) );
			}
			return this;
		}
		if( arguments.sheet.totalColumnCount == 0 )
			return this;
		for( var i=1; i <= arguments.sheet.totalColumnCount; i++ )
			arguments.sheet.columnNames.Append( "column" & i );
		return this;
	}

	private any function forceQueryColumnsToMatchSpecifiedColumns( required struct sheet ){
		if( arguments.sheet.columnNames.Len() >= arguments.sheet.totalColumnCount )
			return this;
		// Not enough columns have been specified. Stash, reset and pad out with defaults
		var specifiedNames = arguments.sheet.columnNames;
		arguments.sheet.columnNames = [];
		cfloop( from=1, to=arguments.sheet.totalColumnCount, index="local.i" ){
			arguments.sheet.columnNames.Append( getQueryColumnNameFromSpecifiedNames( specifiedNames, i ) );
		}
	}

	private string function generateUniqueSheetName( required workbook ){
		var startNumber = ( arguments.workbook.getNumberOfSheets() +1 );
		var maxRetry = ( startNumber +250 );
		for( var sheetNumber = startNumber; sheetNumber <= maxRetry; sheetNumber++ ){
			var proposedName = "Sheet" & sheetNumber;
			if( !sheetExists( arguments.workbook, proposedName ) )
				return proposedName;
		}
		// this should never happen. but if for some reason it did, warn the action failed and abort
		Throw( type=library().getExceptionType(), message="Unable to generate name", detail="Unable to generate a unique sheet name" );
	}

	private numeric function getFirstVisibleSheetNumber( required workbook ){
		var totalSheets = arguments.workbook.getNumberOfSheets();
		cfloop( from=1, to=totalSheets, index="local.sheetNumber" ){
			if( isVisible( arguments.workbook, sheetNumber ) )
				return sheetNumber;
		}
		return 0;
	}

	private string function getQueryColumnNameFromSpecifiedNames( required array specifiedNames, required numeric index ){
		var defaultColumnName = "column" & arguments.index;
		if( arguments.index > arguments.specifiedNames.Len() ) //ACF won't accept IsNull( specifiedNames[ index ] )
			return defaultColumnName;
		var foundColumnName = arguments.specifiedNames[ arguments.index ];
		if( getDataTypeHelper().isString( foundColumnName ) && foundColumnName.Len() )
			return foundColumnName;
		return defaultColumnName;
	}

	private numeric function getSheetIndexFromName( required workbook, required string sheetName ){
		//returns -1 if non-existent
		return arguments.workbook.getSheetIndex( JavaCast( "string", arguments.sheetName ) );
	}

	private string function getVisibility( required workbook, required numeric sheetNumber ){
		validateSheetNumber( arguments.workbook, arguments.sheetNumber );
		var sheetIndex = ( arguments.sheetNumber -1 );
		return arguments.workbook.getSheetVisibility( sheetIndex ).toString();
	}

	private boolean function hasComments( required sheet, required boolean isXlsx ){
		return BooleanFormat( arguments.isXlsx? arguments.sheet.hasComments(): arguments.sheet.getCellComments().Count() );
	}

	private boolean function isActive( required sheet, required boolean isXlsx ){
		if( !arguments.isXlsx )
			return arguments.sheet.isActive();
		var workbook = arguments.sheet.getWorkbook();
		var sheetIndex = workbook.getSheetIndex( arguments.sheet );
		return ( sheetIndex == workbook.getActiveSheetIndex() );
	}

	private boolean function sheetIsEmpty( required sheet ){
		return ( getLastRowIndex( arguments.sheet ) == -1 );
	}

	private boolean function sheetNameArgumentWasProvided(){
		return ( arguments.KeyExists( "sheetName" ) && Len( arguments.sheetName ) );
	}

	private boolean function sheetNumberArgumentWasProvided(){
		return ( arguments.KeyExists( "sheetNumber" ) && Len( arguments.sheetNumber ) );
	}

	private any function throwErrorIFSheetNameAndNumberArgumentsBothMissing(){
		if( !sheetNameArgumentWasProvided( argumentCollection=arguments ) && !sheetNumberArgumentWasProvided( argumentCollection=arguments ) )
			Throw( type=library().getExceptionType(), message="Missing Required Argument", detail="Either sheetName or sheetNumber must be provided" );
		return this;
	}

	private any function throwErrorIFSheetNameAndNumberArgumentsBothPassed(){
		if( sheetNameArgumentWasProvided( argumentCollection=arguments ) && sheetNumberArgumentWasProvided( argumentCollection=arguments ) )
			Throw( type=library().getExceptionType(), message="Invalid arguments", detail="Only one argument is allowed. Specify either a sheetName or sheetNumber, not both" );
		return this;
	}

	private void function doFillMergedCellsWithVisibleValue( required workbook, required sheet ){
		if( !getSheetHelper().hasMergedRegions( arguments.sheet ) )
			return this;
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

}
