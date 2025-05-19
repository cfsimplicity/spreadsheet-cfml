component extends="base"{

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

	numeric function getActiveSheetNumber( required workbook ){
		return arguments.workbook.getActiveSheetIndex()+1;
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

	any function getActiveSheetRowIterator( required workbook ){
		return getActiveSheet( arguments.workbook ).rowIterator();
	}

	array function getAllSheetFormulas( required workbook ){
		var rowIterator = getActiveSheetRowIterator( arguments.workbook );
		var formulas = [];
		while( rowIterator.hasNext() ){
			var cellIterator = rowIterator.next().cellIterator();
			while( cellIterator.hasNext() ){
				var cell = cellIterator.next();
				var cellFormula = getCellFormula( cell );
				if( cellFormula.Len() ) {
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

	string function getPrintOrientation( required workbook, string sheetName, numeric sheetNumber ){
		return getSpecifiedOrActiveSheet( argumentCollection=arguments ).getPrintSetup().getLandscape()? "landscape": "portrait";
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

	numeric function getSheetNumberFromArguments( required workbook, string sheetName, numeric sheetNumber ){
		if( !arguments.KeyExists( "sheetName" ) && !arguments.KeyExists( "sheetNumber" ) )
			return getActiveSheetNumber( arguments.workbook );
		validateSheetNameOrNumberWasProvided( argumentCollection=arguments );
		if( arguments.KeyExists( "sheetName" ) && Len( Trim( arguments.sheetName ) ) ){
			validateSheetExistsWithName( arguments.workbook, arguments.sheetName );
			arguments.sheetNumber = ( arguments.workbook.getSheetIndex( JavaCast( "string", arguments.sheetName ) ) + 1 );
		}
		validateSheetNumber( arguments.workbook, arguments.sheetNumber );
		return arguments.sheetNumber;
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
			,hasDataValidations: BooleanFormat( sheet.getDataValidations().Len() )
			,hasMergedRegions: BooleanFormat( sheet.getNumMergedRegions() )
			,isCurrentActiveSheet: isActive( sheet, isXlsx )
			,isHidden: !isVisible( argumentCollection=arguments )
			,isRightToLeft: sheet.isRightToLeft()
			,name: sheet.getSheetName()
			,numberOfDataValidations: sheet.getDataValidations().Len()
			,numberOfMergedRegions: sheet.getNumMergedRegions()
			,position: getSheetIndexFromName( arguments.workbook, sheet.getSheetName() )+1
			,printOrientation: getPrintOrientation( arguments.workbook, sheet.getSheetName() )
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
		validateSheetNumber( arguments.workbook, arguments.sheetNumber );
		var validStates = [ "HIDDEN", "VERY_HIDDEN", "VISIBLE" ];
		if( !validStates.Find( arguments.visibility ) )
			Throw( type=this.getExceptionType() & ".invalidVisibilityArgument", message="Invalid visibility argument: '#arguments.visibility#'", detail="The visibility must be one of the following: #validStates.ToList( ', ' )#." );
		var visibilityEnum = library().createJavaObject( "org.apache.poi.ss.usermodel.SheetVisibility" )[ JavaCast( "string", arguments.visibility ) ];
		var sheetIndex = ( arguments.sheetNumber -1 );
		arguments.workbook.setSheetVisibility( sheetIndex, visibilityEnum );
		/* POI Docs: "Please note that the sheet currently set as active sheet (sheet 0 in a newly created workbook or the one set via setActiveSheet()) cannot be hidden." */
		if( arguments.visibility == "VISIBLE" || ( arguments.sheetNumber != getActiveSheetNumber( arguments.workbook ) ) )
			return;
		activateFirstVisibleSheet( arguments.workbook );
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

	array function sheetToArrayOfStructs(
		required workbook
		,string sheetName
		,numeric sheetNumber
		,numeric headerRow
		,boolean includeHeaderRow=false
		,boolean includeBlankRows=false
		,boolean includeHiddenColumns=true
		,boolean includeHiddenRows=true
		,boolean fillMergedCellsWithVisibleValue=false
		,boolean includeRichTextFormatting=false
		,string rows //range
		,string columns //range
		,any columnNames //list or array
		,boolean returnVisibleValues=false
	){
		var intermediateResult = sheetToArray( argumentCollection=arguments, forceColumnGeneration=true );
		return ArrayMap( intermediateResult.data, ( row )=>{
			var rowAsOrderedStruct = [:];
			ArrayEach( intermediateResult.columns, ( column, index )=>{
				rowAsOrderedStruct[ column ] = row[ index ];
			});
			return rowAsOrderedStruct;
		})
	}

	struct function sheetToArray(
		required workbook
		,string sheetName
		,numeric sheetNumber
		,numeric headerRow
		,boolean includeHeaderRow=false
		,boolean includeBlankRows=false
		,boolean includeHiddenColumns=true
		,boolean includeHiddenRows=true
		,boolean fillMergedCellsWithVisibleValue=false
		,boolean includeRichTextFormatting=false
		,string rows //range
		,string columns //range
		,any columnNames //list or array
		,boolean returnVisibleValues=false
		,boolean forceColumnGeneration=false
	){
		var result = [ columns: [], data: [] ];//ordered struct
		if( arguments.KeyExists( "sheetName" ) ){
			validateSheetExistsWithName( arguments.workbook, arguments.sheetName );
			arguments.sheetNumber = ( getSheetIndexFromName( arguments.workbook, arguments.sheetName ) +1 );
		}
		else if( !arguments.KeyExists( "sheetNumber" ) )
			arguments.sheetNumber = getFirstVisibleSheetNumber( arguments.workbook );
		if( arguments.sheetNumber == 0 )
			return result;//no visible sheets
		var sheet = readSheet( argumentCollection=arguments );
		if( sheet.hasHeaderRow || sheet.columnNames.Len() || arguments.forceColumnGeneration ){
			generateColumnNames( arguments.workbook, sheet );
			result.columns = sheet.columnNames;
		}
		if( !arguments.includeHiddenColumns && sheet.hasRows ){
			sheet = deleteHiddenColumnsFromArray( sheet );
			if( sheet.totalColumnCount == 0 )
				return result;// all columns were hidden
			if( sheet.columnNames.Len() )
				result.columns = sheet.columnNames;
		}
		result.data = sheet.data;
		return result;
	}

	query function sheetToQuery(
		required workbook
		,string sheetName
		,numeric sheetNumber
		,numeric headerRow
		,boolean includeHeaderRow=false
		,boolean includeBlankRows=false
		,boolean includeHiddenColumns=true
		,boolean includeHiddenRows=true
		,boolean fillMergedCellsWithVisibleValue=false
		,boolean includeRichTextFormatting=false
		,string rows //range
		,string columns //range
		,any columnNames //list or array
		,any queryColumnTypes="" //'auto', single default type e.g. 'VARCHAR', or list of types, or struct of column names/types mapping. Empty means no types are specified.
		,boolean makeColumnNamesSafe=false
		,boolean returnVisibleValues=false
	){
		if( arguments.KeyExists( "sheetName" ) ){
			validateSheetExistsWithName( arguments.workbook, arguments.sheetName );
			arguments.sheetNumber = ( getSheetIndexFromName( arguments.workbook, arguments.sheetName ) +1 );
		}
		else if( !arguments.KeyExists( "sheetNumber" ) )
			arguments.sheetNumber = getFirstVisibleSheetNumber( arguments.workbook );
		if( arguments.sheetNumber == 0 )
			return QueryNew( "" );//no visible sheets
		var sheet = readSheet( argumentCollection=arguments );
		generateColumnNames( arguments.workbook, sheet );
		arguments.queryColumnTypes = getQueryHelper().parseQueryColumnTypesArgument( arguments.queryColumnTypes, sheet.columnNames, sheet.totalColumnCount, sheet.data );
		var result = getQueryHelper()._QueryNew( sheet.columnNames, arguments.queryColumnTypes, sheet.data, arguments.makeColumnNamesSafe );
		if( !arguments.includeHiddenColumns && sheet.hasRows ){
			result = getQueryHelper().deleteHiddenColumnsFromQuery( sheet, result );
			if( sheet.totalColumnCount == 0 )
				return QueryNew( "" );// all columns were hidden: return a blank query.
		}
		return result;
	}

	any function validateSheetExistsWithName( required workbook, required string sheetName ){
		if( !sheetExists( workbook=arguments.workbook, sheetName=arguments.sheetName ) )
			Throw( type=library().getExceptionType() & ".invalidSheetName", message="Invalid sheet name [#arguments.sheetName#]", detail="The specified sheet was not found in the current workbook." );
		return this;
	}

	any function validateSheetNumber( required workbook, required numeric sheetNumber ){
		if( !sheetExists( workbook=arguments.workbook, sheetNumber=arguments.sheetNumber ) ){
			var sheetCount = arguments.workbook.getNumberOfSheets();
			Throw( type=library().getExceptionType() & ".invalidSheetNumber", message="Invalid sheet number [#arguments.sheetNumber#]", detail="The sheetNumber must a whole number between 1 and the total number of sheets in the workbook [#sheetCount#]" );
		}
		return this;
	}

	any function validateSheetName( required string sheetName ){
		var characterCount = Len( arguments.sheetName );
		if( characterCount > 31 )
			Throw( type=library().getExceptionType() & ".invalidSheetName", message="Invalid sheet name", detail="The sheetname contains too many characters [#characterCount#]. The maximum is 31." );
		var poiTool = library().createJavaObject( "org.apache.poi.ss.util.WorkbookUtil" );
		try{
			poiTool.validateSheetName( JavaCast( "String", arguments.sheetName ) );
		}
		catch( "java.lang.IllegalArgumentException" exception ){
			Throw( type=library().getExceptionType() & ".invalidCharacters", message="Invalid characters in sheet name", detail=exception.message );
		}
		catch( "java.lang.reflect.InvocationTargetException" exception ){
			//ACF
			Throw( type=library().getExceptionType() & ".invalidCharacters", message="Invalid characters in sheet name", detail=exception.message );
		}
		return this;
	}

	void function validateSheetNameOrNumberWasProvided(){
		throwErrorIFSheetNameAndNumberArgumentsBothMissing( argumentCollection=arguments );
		throwErrorIFSheetNameAndNumberArgumentsBothPassed( argumentCollection=arguments );
	}

	any function throwErrorIFSheetNameAndNumberArgumentsBothPassed(){
		if( sheetNameArgumentWasProvided( argumentCollection=arguments ) && sheetNumberArgumentWasProvided( argumentCollection=arguments ) )
			Throw( type=library().getExceptionType() & ".invalidArguments", message="Invalid arguments", detail="Specify either a sheetName or sheetNumber, not both" );
		return this;
	}

	/* Private */

	private string function getCellFormula( required cell ) {
		if( getCellHelper().cellIsOfType( cell, "FORMULA" ) )
			return cell.getCellFormula();
		return "";
	}

	private any function generateColumnNames( required workbook, required struct sheet ){
		if( arguments.sheet.columnNames.Len() ){
			forceColumnsToMatchSpecifiedColumns( arguments.sheet );
			return this; // already generated
		}
		if( sheetIsEmpty( arguments.sheet.object, arguments.workbook ) )
			return this;
		if( arguments.sheet.hasHeaderRow ){
			// use specified header row values as column names
			var headerRowObject = getRowHelper().getRowFromSheet( arguments.workbook, arguments.sheet.object, arguments.sheet.headerRowIndex );
			var headerRowData = getRowHelper().getRowData( arguments.workbook, headerRowObject, arguments.sheet.columnRanges );
			// adds default column names if header row column count is less than total data column count
			cfloop( from=1, to=arguments.sheet.totalColumnCount, index="local.i" ){
				arguments.sheet.columnNames.Append( getColumnNameFromSpecifiedNames( headerRowData, i ) );
			}
			return this;
		}
		if( arguments.sheet.totalColumnCount == 0 )
			return this;
		for( var i=1; i <= arguments.sheet.totalColumnCount; i++ )
			arguments.sheet.columnNames.Append( "column" & i );
		return this;
	}

	private any function forceColumnsToMatchSpecifiedColumns( required struct sheet ){
		if( arguments.sheet.columnNames.Len() >= arguments.sheet.totalColumnCount )
			return this;
		// Not enough columns have been specified. Stash, reset and pad out with defaults
		var specifiedNames = arguments.sheet.columnNames;
		arguments.sheet.columnNames = [];
		cfloop( from=1, to=arguments.sheet.totalColumnCount, index="local.i" ){
			arguments.sheet.columnNames.Append( getColumnNameFromSpecifiedNames( specifiedNames, i ) );
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
		Throw( type=library().getExceptionType() & ".uniqueNameGenerationFailure", message="Unable to generate name", detail="Unable to generate a unique sheet name" );
	}

	private numeric function getFirstVisibleSheetNumber( required workbook ){
		var totalSheets = arguments.workbook.getNumberOfSheets();
		cfloop( from=1, to=totalSheets, index="local.sheetNumber" ){
			if( isVisible( arguments.workbook, sheetNumber ) )
				return sheetNumber;
		}
		return 0;
	}

	private string function getColumnNameFromSpecifiedNames( required array specifiedNames, required numeric index ){
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
		if( getStreamingReaderHelper().isStreamingReaderFormat( arguments.workbook ) ) // getSheetVisibility() not supported
			return "VISIBLE";
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

	private boolean function sheetIsEmpty( required sheet, required workbook ){
		return !arguments.sheet.rowIterator().hasNext();
	}

	private boolean function sheetNameArgumentWasProvided(){
		return ( arguments.KeyExists( "sheetName" ) && Len( arguments.sheetName ) );
	}

	private boolean function sheetNumberArgumentWasProvided(){
		return ( arguments.KeyExists( "sheetNumber" ) && Len( arguments.sheetNumber ) );
	}

	private any function throwErrorIFSheetNameAndNumberArgumentsBothMissing(){
		if( !sheetNameArgumentWasProvided( argumentCollection=arguments ) && !sheetNumberArgumentWasProvided( argumentCollection=arguments ) )
			Throw( type=library().getExceptionType() & ".missingRequiredArgument", message="Missing Required Argument", detail="Either sheetName or sheetNumber must be provided" );
		return this;
	}

	private void function doFillMergedCellsWithVisibleValue( required workbook, required sheet ){
		if( !hasMergedRegions( arguments.sheet ) )
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

	private void function populateSheetData(
		required workbook
		,required sheet
		,required boolean fillMergedCellsWithVisibleValue
		,required boolean includeRichTextFormatting
		,required boolean returnVisibleValues
	){
		var addRowToSheetDataArgs = {
			workbook: arguments.workbook
			,sheet: arguments.sheet
			,includeRichTextFormatting: arguments.includeRichTextFormatting
			,returnVisibleValues: arguments.returnVisibleValues
		};
		if( getStreamingReaderHelper().isStreamingReaderFormat( arguments.workbook ) ){
			var rowIterator = arguments.sheet.object.rowIterator();
			while( rowIterator.hasNext() ){
				addRowToSheetDataArgs.rowObject = rowIterator.next();
				addRowToSheetDataArgs.rowIndex = addRowToSheetDataArgs.rowObject.getRowNum();
				getRowHelper().addRowToSheetData( argumentCollection=addRowToSheetDataArgs );
			}
			return;
		}
		if( arguments.fillMergedCellsWithVisibleValue )
			doFillMergedCellsWithVisibleValue( arguments.workbook, arguments.sheet.object );
		if( arguments.KeyExists( "rows" ) ){
			var allRanges = getRangeHelper().extractRanges( arguments.rows, arguments.workbook );
			for( var thisRange in allRanges ){
				for( var rowNumber = thisRange.startAt; rowNumber <= thisRange.endAt; rowNumber++ ){
					addRowToSheetDataArgs.rowIndex = ( rowNumber -1 );
					getRowHelper().addRowToSheetData( argumentCollection=addRowToSheetDataArgs );
				}
			}
		}
		else{
			var lastRowIndex = arguments.sheet.object.getLastRowNum();// zero based
			for( var rowIndex = 0; rowIndex <= lastRowIndex; rowIndex++ ){
				addRowToSheetDataArgs.rowIndex = rowIndex;
				getRowHelper().addRowToSheetData( argumentCollection=addRowToSheetDataArgs );
			}
		}
	}

	private void function activateFirstVisibleSheet( required workbook ){
		var firstVisibleSheetNumber = getFirstVisibleSheetNumber( arguments.workbook );
		if( firstVisibleSheetNumber == 0  ) // there are no visible sheets
			return;
		library().setActiveSheetNumber( arguments.workbook, firstVisibleSheetNumber );
	}

	private struct function readSheet(
		required workbook
		,numeric sheetNumber
		,numeric headerRow
		,boolean includeHeaderRow=false
		,boolean includeBlankRows=false
		,boolean includeHiddenRows=true
		,boolean fillMergedCellsWithVisibleValue=false
		,boolean includeRichTextFormatting=false
		,string rows //range
		,string columns //range
		,any columnNames //list or array
		,boolean returnVisibleValues=false
	){
		var sheet = {
			includeHeaderRow: arguments.includeHeaderRow
			,hasHeaderRow: ( arguments.KeyExists( "headerRow" ) && Val( arguments.headerRow ) )
			,includeBlankRows: arguments.includeBlankRows
			,includeHiddenRows: arguments.includeHiddenRows
			,columnNames: []
			,columnRanges: []
			,totalColumnCount: 0
			,data: []
			,hasRows: false
		};
		if( arguments.KeyExists( "columnNames" ) && arguments.columnNames.Len() )
			sheet.columnNames = IsArray( arguments.columnNames )? arguments.columnNames: arguments.columnNames.ListToArray();
		sheet.headerRowIndex = sheet.hasHeaderRow? ( arguments.headerRow -1 ): -1;
		if( arguments.KeyExists( "columns" ) ){
			sheet.columnRanges = getRangeHelper().extractRanges( arguments.columns, arguments.workbook, "column" );
			sheet.totalColumnCount = getColumnHelper().columnCountFromRanges( sheet.columnRanges );
		}
		sheet.object = getSheetByNumber( arguments.workbook, arguments.sheetNumber );
		sheet.hasRows = !sheetIsEmpty( sheet.object, arguments.workbook );
		if( sheet.hasRows  ){
			var populateDataArgs = {
				workbook: arguments.workbook
				,sheet: sheet
				,fillMergedCellsWithVisibleValue: arguments.fillMergedCellsWithVisibleValue
				,includeRichTextFormatting: arguments.includeRichTextFormatting
				,returnVisibleValues: arguments.returnVisibleValues
			};
			if( arguments.KeyExists( "rows" ) )
				populateDataArgs.rows = arguments.rows;
			populateSheetData( argumentCollection=populateDataArgs );
		}
		return sheet;
	}

	private struct function deleteHiddenColumnsFromArray( required sheet ){
		var startIndex = ( arguments.sheet.totalColumnCount -1 );
		for( var colIndex = startIndex; colIndex >= 0; colIndex-- ){
			if( !arguments.sheet.object.isColumnHidden( JavaCast( "int", colIndex ) ) )
				continue;
			var columnNumber = ( colIndex +1 );
			arguments.sheet.data = _ArrayDeleteColumn( arguments.sheet.data, columnNumber );
			arguments.sheet.totalColumnCount--;
			if( arguments.sheet.columnNames.Len() )
				arguments.sheet.columnNames = _ArrayDeleteAt( arguments.sheet.columnNames, columnNumber );
		}
		return arguments.sheet;
	}

	private array function _ArrayDeleteColumn( required array data, required numeric position ){
		cfloop( array=arguments.data, item="local.row", index="local.i" ){
			ArrayDeleteAt( row, arguments.position );
			if( library().getIsACF() )
				arguments.data[ i ] = row; //ACF doesn't replace the row
		}
		return arguments.data;
	}

	private array function _ArrayDeleteAt( required array data, required numeric position ){
		// return the resulting array rather than a boolean flag
		ArrayDeleteAt( arguments.data, arguments.position );
		return arguments.data;
	}

}