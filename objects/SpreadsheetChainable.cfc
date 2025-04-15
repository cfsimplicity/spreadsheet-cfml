component{

	property name="library";
	property name="workbook";

	public SpreadsheetChainable function init( required spreadsheetLibrary, required existingWorkbookOrNewWorkbookType ){
		variables.library = arguments.spreadsheetLibrary;
		setupWorkbook( arguments.existingWorkbookOrNewWorkbookType );
		return this;
	}

	public any function getWorkbook(){
		return variables.workbook;
	}

	private void function setupWorkbook( required existingWorkbookOrNewWorkbookType ){
		if( variables.library.isSpreadsheetObject( arguments[ 1 ] ) ){
			variables.workbook = arguments[ 1 ];
			return;
		}
		var newWorkbookType = arguments[ 1 ];
		if( IsSimpleValue( newWorkbookType ) ){
			switch( newWorkbookType ){
				case "xls": variables.workbook = variables.library.newXls();
					return;
				case "xlsx": variables.workbook = variables.library.newXlsx();
					return;
				case "streamingXlsx": case "streamingXml":
					variables.workbook = variables.library.newStreamingXlsx();
					return;
				case "": //allowed so workbook can be read post-init()
					return;
			}
		}
		// anything else is not valid
		throwErrorIfWorkbookIsInvalid();
	}

	private void function addWorkbookArgument( required args ){
		throwErrorIfWorkbookIsNull();
		arguments.args.workbook = variables.workbook;
	}

	private void function throwErrorIfWorkbookIsNull(){
		if( IsNull( variables.workbook ) )
			Throw( type=variables.library.getExceptionType() & ".missingWorkbook", message="Missing workbook", detail="No workbook object has been specified for this chained call. You can specify a new workbook type or pass in an existing object on initialisation, or read a file, query or csv on the second call after initialisation" );
	}

	private void function throwErrorIfWorkbookIsInvalid(){
		if( !variables.library.isSpreadsheetObject( variables.workbook?:"" ) )
			Throw( type=variables.library.getExceptionType() & ".invalidWorkbook", message="Invalid workbook", detail="The workbook specified in the chained call is not a valid spreadsheet object" );
	}

	/* PUBLIC API */
	
	public SpreadsheetChainable function addAutofilter( string cellRange="", numeric row=1 ){
		addWorkbookArgument( arguments );
		variables.library.addAutofilter( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function addColumn(
		required data // Delimited list of values OR array
		,numeric startRow
		,numeric startColumn
		,boolean insert=false
		,string delimiter=","
		,boolean autoSize=false
	){
		addWorkbookArgument( arguments );
		variables.library.addColumn( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function addConditionalFormatting( required ConditionalFormatting conditionalFormatting ){
		addWorkbookArgument( arguments );
		variables.library.addConditionalFormatting( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function addDataValidation( required DataValidation dataValidation ){
		addWorkbookArgument( arguments );
		variables.library.addDataValidation( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function addFreezePane(
		required numeric freezeColumn
		,required numeric freezeRow
		,numeric leftmostColumn //left column visible in right pane
		,numeric topRow //top row visible in bottom pane
	){
		addWorkbookArgument( arguments );
		variables.library.addFreezePane( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function addImage(
		string filepath
		,imageData
		,string imageType
		,required string anchor
	){
		addWorkbookArgument( arguments );
		variables.library.addImage( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function addInfo( required struct info ){
		addWorkbookArgument( arguments );
		variables.library.addInfo( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function addPageBreaks( string rowBreaks="", string columnBreaks="" ){
		addWorkbookArgument( arguments );
		variables.library.addPageBreaks( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function addPrintGridLines(){
		addWorkbookArgument( arguments );
		variables.library.addPrintGridLines( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function addRow(
		required data // Delimited list of data, OR array
		,numeric row
		,numeric column=1
		,boolean insert=true
		,string delimiter=","
		,boolean handleEmbeddedCommas=true // When true, values enclosed in single quotes are treated as a single element like in ACF. Only applies when the delimiter is a comma.
		,boolean autoSizeColumns=false
		,struct datatypes
	){
		addWorkbookArgument( arguments );
		variables.library.addRow( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function addRows(
		required data // query or array
		,numeric row
		,numeric column=1
		,boolean insert=true
		,boolean autoSizeColumns=false
		,boolean includeQueryColumnNames=false
		,boolean ignoreQueryColumnDataTypes=false
		,struct datatypes
	){
		addWorkbookArgument( arguments );
		variables.library.addRows( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function addSplitPane(
		required numeric xSplitPosition
		,required numeric ySplitPosition
		,required numeric leftmostColumn
		,required numeric topRow
		,string activePane="UPPER_LEFT" //Valid values are LOWER_LEFT, LOWER_RIGHT, UPPER_LEFT, and UPPER_RIGHT
	){
		addWorkbookArgument( arguments );
		variables.library.addSplitPane( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function autoSizeColumn( required numeric column, boolean useMergedCells=false ){
		addWorkbookArgument( arguments );
		variables.library.autoSizeColumn( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function cleanUpStreamingXml(){
		addWorkbookArgument( arguments );
		variables.library.cleanUpStreamingXml( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function clearCell( required numeric row, required numeric column ){
		addWorkbookArgument( arguments );
		variables.library.clearCell( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function clearCellRange(
		required numeric startRow
		,required numeric startColumn
		,required numeric endRow
		,required numeric endColumn
	){
		addWorkbookArgument( arguments );
		variables.library.clearCellRange( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function collapseColumnGroup( required numeric column ){
		addWorkbookArgument( arguments );
		variables.library.collapseColumnGroup( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function collapseRowGroup( required numeric row ){
		addWorkbookArgument( arguments );
		variables.library.collapseRowGroup( argumentCollection=arguments );
		return this;
	}

	// Ends chain - returns CellStyle object
	public any function createCellStyle( required struct format ){
		addWorkbookArgument( arguments );
		return variables.library.createCellStyle( argumentCollection=arguments );
	}

	public SpreadsheetChainable function createSheet( string sheetName, overwrite=false ){
		addWorkbookArgument( arguments );
		variables.library.createSheet( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function deleteColumn( required numeric column ){
		addWorkbookArgument( arguments );
		variables.library.deleteColumn( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function deleteColumns( required string range ){
		addWorkbookArgument( arguments );
		variables.library.deleteColumns( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function deleteRow( required numeric row ){
		addWorkbookArgument( arguments );
		variables.library.deleteRow( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function deleteRows( required string range ){
		addWorkbookArgument( arguments );
		variables.library.deleteRows( argumentCollection=arguments );
		return this;
	}

	// Ends chain
	public void function download( required string filename, string contentType ){
		addWorkbookArgument( arguments );
		variables.library.download( argumentCollection=arguments );
	}

	public SpreadsheetChainable function expandColumnGroup( required numeric column ){
		addWorkbookArgument( arguments );
		variables.library.expandColumnGroup( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function expandRowGroup( required numeric row ){
		addWorkbookArgument( arguments );
		variables.library.expandRowGroup( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function formatCell(
		struct format={}
		,required numeric row
		,required numeric column
		,boolean overwriteCurrentStyle=true
		,any cellStyle
	){
		addWorkbookArgument( arguments );
		variables.library.formatCell( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function formatCellRange(
		struct format={}
		,required numeric startRow
		,required numeric endRow
		,required numeric startColumn
		,required numeric endColumn
		,boolean overwriteCurrentStyle=true
		,any cellStyle
	){
		addWorkbookArgument( arguments );
		variables.library.formatCellRange( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function formatColumn(
		struct format={}
		,required numeric column
		,boolean overwriteCurrentStyle=true
		,any cellStyle
	){
		addWorkbookArgument( arguments );
		variables.library.formatColumn( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function formatColumns(
		struct format={}
		,required string range
		,boolean overwriteCurrentStyle=true
		,any cellStyle
	){
		addWorkbookArgument( arguments );
		variables.library.formatColumns( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function formatRow(
		struct format={}
		,required numeric row
		,boolean overwriteCurrentStyle=true
		,any cellStyle
	){
		addWorkbookArgument( arguments );
		variables.library.formatRow( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function formatRows(
		struct format={}
		,required string range
		,boolean overwriteCurrentStyle=true
		,any cellStyle
	){
		addWorkbookArgument( arguments );
		variables.library.formatRows( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function fromCsv(){
		variables.workbook = variables.library.workbookFromCsv( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function fromQuery(){
		variables.workbook = variables.library.workbookFromQuery( argumentCollection=arguments );
		return this;
	}

	// Ends chain
	public struct function getActiveCell(){
		addWorkbookArgument( arguments );
		return variables.library.getActiveCell( argumentCollection=arguments );
	}

	// Ends chain
	public string function getCellAddress( required numeric row, required numeric column ){
		addWorkbookArgument( arguments );
		return variables.library.getCellAddress( argumentCollection=arguments );
	}

	public any function getCellComment( numeric row, numeric column ){
		// Ends chain: returns struct OR array of structs
		addWorkbookArgument( arguments );
		return variables.library.getCellComment( argumentCollection=arguments );
	}

	//Ends chain
	public array function getCellComments(){
		addWorkbookArgument( arguments );
		return variables.library.getCellComments( argumentCollection=arguments );
	}

	//Ends chain
	public struct function getCellFormat( required numeric row, required numeric column ){
		addWorkbookArgument( arguments );
		return variables.library.getCellFormat( argumentCollection=arguments );
	}

	public any function getCellFormula( numeric row, numeric column ){
		// Ends chain: returns string OR array of strings
		addWorkbookArgument( arguments );
		return variables.library.getCellFormula( argumentCollection=arguments );
	}

	// Ends chain
	public string function getCellHyperLink( required numeric row, required numeric column ){
		addWorkbookArgument( arguments );
		return variables.library.getCellHyperLink( argumentCollection=arguments );
	}

	// Ends chain
	public string function getCellType( required numeric row, required numeric column ){
		addWorkbookArgument( arguments );
		return variables.library.getCellType( argumentCollection=arguments );
	}

	// Ends chain
	public any function getCellValue( required numeric row, required numeric column, boolean returnVisibleValue=true ){
		addWorkbookArgument( arguments );
		return variables.library.getCellValue( argumentCollection=arguments );
	}

	// Ends chain
	public numeric function getColumnCount( sheetNameOrNumber ){
		addWorkbookArgument( arguments );
		return variables.library.getColumnCount( argumentCollection=arguments );
	}

	// Ends chain
	public numeric function getColumnWidth( required numeric column ){
		addWorkbookArgument( arguments );
		return variables.library.getColumnWidth( argumentCollection=arguments );
	}

	// Ends chain
	public numeric function getColumnWidthInPixels( required numeric column ){
		addWorkbookArgument( arguments );
		return variables.library.getColumnWidthInPixels( argumentCollection=arguments );
	}

	// Ends chain
	public numeric function getLastRowNumber( sheetNameOrNumber ){
		addWorkbookArgument( arguments );
		return variables.library.getLastRowNumber( argumentCollection=arguments );
	}

	// Ends chain
	public boolean function getRecalculateFormulasOnNextOpen( string sheetName ){
		addWorkbookArgument( arguments );
		return variables.library.getRecalculateFormulasOnNextOpen( argumentCollection=arguments );
	}

	// Ends chain
	public numeric function getRowCount( sheetNameOrNumber ){
		addWorkbookArgument( arguments );
		return variables.library.getRowCount( argumentCollection=arguments );
	}

	// Ends chain
	public string function getSheetPrintOrientation( string sheetName, numeric sheetNumber ){
		addWorkbookArgument( arguments );
		return variables.library.getSheetPrintOrientation( argumentCollection=arguments );
	}

	public SpreadsheetChainable function groupColumns( required numeric startColumn, required numeric endColumn ){
		addWorkbookArgument( arguments );
		variables.library.groupColumns( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function groupRows( required numeric startRow, required numeric endRow ){
		addWorkbookArgument( arguments );
		variables.library.groupRows( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function hideColumn( required numeric column ){
		addWorkbookArgument( arguments );
		variables.library.hideColumn( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function hideRow( required numeric row ){
		addWorkbookArgument( arguments );
		variables.library.hideRow( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function hideSheet( string sheetName, numeric sheetNumber ){
		addWorkbookArgument( arguments );
		variables.library.hideSheet( argumentCollection=arguments );
		return this;
	}

	// Ends chain
	public struct function info(){
		// argument name is workbookOrPath not workbook so custom handling
		throwErrorIfWorkbookIsInvalid();
		return variables.library.info( workbookOrPath=variables.workbook );
	}

	public SpreadsheetChainable function mergeCells(
		required numeric startRow
		,required numeric endRow
		,required numeric startColumn
		,required numeric endColumn
		,boolean emptyInvisibleCells=false
	){
		addWorkbookArgument( arguments );
		variables.library.mergeCells( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function moveSheet( required string sheetName, required numeric newPosition ){
		addWorkbookArgument( arguments );
		variables.library.moveSheet( argumentCollection=arguments );
		return this;
	}

	public any function read(
		required string src
		,string format
		,string columns
		,any columnNames //list or array
		,numeric headerRow
		,string rows
		,string sheetName
		,numeric sheetNumber // 1-based
		,boolean includeHeaderRow=false
		,boolean includeBlankRows=false
		,boolean fillMergedCellsWithVisibleValue=false
		,boolean includeHiddenColumns=true
		,boolean includeHiddenRows=true
		,boolean includeRichTextFormatting=false
		,string password
		,string csvDelimiter=","
		,any queryColumnTypes //'auto', list of types, or struct of column names/types mapping. Null means no types are specified.
		,boolean makeColumnNamesSafe=false
		,boolean returnVisibleValues=false
	){
		if( arguments.KeyExists( "format" ) )
			return variables.library.read( argumentCollection=arguments );
		variables.workbook = variables.library.read( argumentCollection=arguments );
		return this;
	}

	// Ends chain
	public binary function readBinary(){
		addWorkbookArgument( arguments );
		return variables.library.readBinary( argumentCollection=arguments );
	}

	public SpreadsheetChainable function recalculateAllFormulas(){
		addWorkbookArgument( arguments );
		variables.library.recalculateAllFormulas( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function removeColumnBreak( required numeric column ){
		addWorkbookArgument( arguments );
		variables.library.removeColumnBreak( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function removeRowBreak( required numeric row ){
		addWorkbookArgument( arguments );
		variables.library.removeRowBreak( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function removePrintGridlines(){
		addWorkbookArgument( arguments );
		variables.library.removePrintGridlines( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function removeSheet( required string sheetName ){
		addWorkbookArgument( arguments );
		variables.library.removeSheet( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function removeSheetNumber( required numeric sheetNumber ){
		addWorkbookArgument( arguments );
		variables.library.removeSheetNumber( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function renameSheet( required string sheetName, required numeric sheetNumber ){
		addWorkbookArgument( arguments );
		variables.library.renameSheet( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setActiveCell( required numeric row, required numeric column ){
		addWorkbookArgument( arguments );
		variables.library.setActiveCell( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setActiveSheet( string sheetName, numeric sheetNumber ){
		addWorkbookArgument( arguments );
		variables.library.setActiveSheet( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setActiveSheetNumber( numeric sheetNumber ){
		addWorkbookArgument( arguments );
		variables.library.setActiveSheetNumber( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setCellComment(
		required struct comment
		,required numeric row
		,required numeric column
	){
		addWorkbookArgument( arguments );
		variables.library.setCellComment( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setCellFormula(
		required string formula
		,required numeric row
		,required numeric column
	){
		addWorkbookArgument( arguments );
		variables.library.setCellFormula( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setCellHyperlink(
		required string link
		,required numeric row
		,required numeric column
		,any cellValue
		,string type="URL"
		,struct format={ color: "BLUE", underline: true }
		,string tooltip //xlsx only, maybe MS Excel full version only
	){
		addWorkbookArgument( arguments );
		variables.library.setCellHyperlink( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setCellRangeValue(
		required value
		,required numeric startRow
		,required numeric endRow
		,required numeric startColumn
		,required numeric endColumn
	){
		addWorkbookArgument( arguments );
		variables.library.setCellRangeValue( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setCellValue( required value, required numeric row, required numeric column, string datatype ){
		addWorkbookArgument( arguments );
		variables.library.setCellValue( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setColumnWidth( required numeric column, required numeric width ){
		addWorkbookArgument( arguments );
		variables.library.setColumnWidth( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setFitToPage( required boolean state, numeric pagesWide, numeric pagesHigh ){
		addWorkbookArgument( arguments );
		variables.library.setFitToPage( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setFooter(
		string leftFooter=""
		,string centerFooter=""
		,string rightFooter=""
	){
		addWorkbookArgument( arguments );
		variables.library.setFooter( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setFooterImage(
		required string position // left|center|right
		,required any image
		,string imageType
	){
		addWorkbookArgument( arguments );
		variables.library.setFooterImage( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setHeader(
		string leftHeader=""
		,string centerHeader=""
		,string rightHeader=""
	){
		addWorkbookArgument( arguments );
		variables.library.setHeader( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setHeaderImage(
		required string position // left|center|right
		,required any image
		,string imageType
	){
		addWorkbookArgument( arguments );
		variables.library.setHeaderImage( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setReadOnly( required string password ){
		addWorkbookArgument( arguments );
		variables.library.setReadOnly( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setRecalculateFormulasOnNextOpen( boolean value=true, string sheetName ){
		addWorkbookArgument( arguments );
		variables.library.setRecalculateFormulasOnNextOpen( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setRepeatingColumns( required string columnRange ){
		addWorkbookArgument( arguments );
		variables.library.setRepeatingColumns( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setRepeatingRows( required string rowRange ){
		addWorkbookArgument( arguments );
		variables.library.setRepeatingRows( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setColumnBreak( required numeric column ){
		addWorkbookArgument( arguments );
		variables.library.setColumnBreak( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setRowBreak( required numeric row ){
		addWorkbookArgument( arguments );
		variables.library.setRowBreak( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setRowHeight( required numeric row, required numeric height ){
		addWorkbookArgument( arguments );
		variables.library.setRowHeight( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setSheetTopMargin( required numeric marginSize, string sheetName, numeric sheetNumber ){
		addWorkbookArgument( arguments );
		variables.library.setSheetTopMargin( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setSheetBottomMargin( required numeric marginSize, string sheetName, numeric sheetNumber ){
		addWorkbookArgument( arguments );
		variables.library.setSheetBottomMargin( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setSheetLeftMargin( required numeric marginSize, string sheetName, numeric sheetNumber ){
		addWorkbookArgument( arguments );
		variables.library.setSheetLeftMargin( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setSheetRightMargin( required numeric marginSize, string sheetName, numeric sheetNumber ){
		addWorkbookArgument( arguments );
		variables.library.setSheetRightMargin( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setSheetHeaderMargin( required numeric marginSize, string sheetName, numeric sheetNumber ){
		addWorkbookArgument( arguments );
		variables.library.setSheetHeaderMargin( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setSheetFooterMargin( required numeric marginSize, string sheetName, numeric sheetNumber ){
		addWorkbookArgument( arguments );
		variables.library.setSheetFooterMargin( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setSheetPrintOrientation( required string mode, string sheetName, numeric sheetNumber ){
		addWorkbookArgument( arguments );
		variables.library.setSheetPrintOrientation( argumentCollection=arguments );
		return this;
	}

	// Ends chain
	public struct function sheetInfo( numeric sheetNumber ){
		addWorkbookArgument( arguments );
		return variables.library.sheetInfo( argumentCollection=arguments );
	}

	public SpreadsheetChainable function shiftColumns( required numeric start, numeric end=arguments.start, numeric offset=1 ){
		addWorkbookArgument( arguments );
		variables.library.shiftColumns( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function shiftRows( required numeric start, numeric end=arguments.start, numeric offset=1 ){
		addWorkbookArgument( arguments );
		variables.library.shiftRows( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function showColumn( required numeric column ){
		addWorkbookArgument( arguments );
		variables.library.showColumn( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function showRow( required numeric row ){
		addWorkbookArgument( arguments );
		variables.library.showRow( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function ungroupColumns( required numeric startColumn, required numeric endColumn ){
		addWorkbookArgument( arguments );
		variables.library.ungroupColumns( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function ungroupRows( required numeric startRow, required numeric endRow ){
		addWorkbookArgument( arguments );
		variables.library.ungroupRows( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function unhideSheet( string sheetName, numeric sheetNumber ){
		addWorkbookArgument( arguments );
		variables.library.unhideSheet( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function write(
		required string filepath
		,boolean overwrite=false
		,string password
		,string algorithm="agile"
	){
		addWorkbookArgument( arguments );
		variables.library.write( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function writeToCsv(
		required string filepath
		,boolean overwrite=false
		,string delimiter=","
		,boolean includeHeaderRow=true
		,numeric headerRow=1
	){
		addWorkbookArgument( arguments );
		variables.library.writeToCsv( argumentCollection=arguments );
		return this;
	}

}