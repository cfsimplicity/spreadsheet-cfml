component accessors="true"{

	property name="library";
	property name="workbook";

	public SpreadsheetChainable function init( required spreadsheetLibrary, required existingWorkbookOrNewWorkbookType ){
		this.setLibrary( arguments.spreadsheetLibrary );
		setupWorkbook( arguments.existingWorkbookOrNewWorkbookType );
		return this;
	}

	private void function setupWorkbook( required existingWorkbookOrNewWorkbookType ){
		if( this.getLibrary().isSpreadsheetObject( arguments.existingWorkbookOrNewWorkbookType ) ){
			var workbook = arguments.existingWorkbookOrNewWorkbookType;
			this.setWorkbook( workbook );
			return;
		}
		var newWorkbookType = arguments.existingWorkbookOrNewWorkbookType;
		switch( newWorkbookType ){
			case "xls": this.setWorkbook( this.getLibrary().newXls() );
				return;
			case "xlsx": this.setWorkbook( this.getLibrary().newXlsx() );
				return;
			case "streamingXlsx": case "streamingXml":
				this.setWorkbook( this.getLibrary().newStreamingXlsx() );
		}
	}

	private void function addWorkbookArgument( required args ){
		throwErrorIfWorkbookIsNull();
		throwErrorIfWorkbookIsInvalid();
		arguments.args.workbook = this.getWorkbook();
	}

	private void function throwErrorIfWorkbookIsNull(){
		if( IsNull( this.getWorkbook() ) )
			Throw( type=this.getLibrary().getExceptionType(), message="Missing workbook", detail="No workbook object has been specified for this chained call. You can specify a new workbook type or pass in an existing object on initialisation, or read a file, query or csv on the second call after initialisation" );
	}

	private void function throwErrorIfWorkbookIsInvalid(){
		if( !this.getLibrary().isSpreadsheetObject( this.getWorkbook() ) )
			Throw( type=this.getLibrary().getExceptionType(), message="Invalid workbook", detail="The workbook specified in the chained call is not a valid spreadsheet object" );
	}

	/* PUBLIC API */
	
	public SpreadsheetChainable function addAutofilter( string cellRange="", numeric row=1 ){
		addWorkbookArgument( arguments );
		this.getLibrary().addAutofilter( argumentCollection=arguments );
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
		this.getLibrary().addColumn( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function addFreezePane(
		required numeric freezeColumn
		,required numeric freezeRow
		,numeric leftmostColumn //left column visible in right pane
		,numeric topRow //top row visible in bottom pane
	){
		addWorkbookArgument( arguments );
		this.getLibrary().addFreezePane( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function addImage(
		string filepath
		,imageData
		,string imageType
		,required string anchor
	){
		addWorkbookArgument( arguments );
		this.getLibrary().addImage( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function addInfo( required struct info ){
		addWorkbookArgument( arguments );
		this.getLibrary().addInfo( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function addPageBreaks( string rowBreaks="", string columnBreaks="" ){
		addWorkbookArgument( arguments );
		this.getLibrary().addPageBreaks( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function addPrintGridLines(){
		addWorkbookArgument( arguments );
		this.getLibrary().addPrintGridLines( argumentCollection=arguments );
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
		this.getLibrary().addRow( argumentCollection=arguments );
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
		this.getLibrary().addRows( argumentCollection=arguments );
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
		this.getLibrary().addSplitPane( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function autoSizeColumn( required numeric column, boolean useMergedCells=false ){
		addWorkbookArgument( arguments );
		this.getLibrary().autoSizeColumn( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function cleanUpStreamingXml(){
		addWorkbookArgument( arguments );
		this.getLibrary().cleanUpStreamingXml( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function clearCell( required numeric row, required numeric column ){
		addWorkbookArgument( arguments );
		this.getLibrary().clearCell( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function clearCellRange(
		required numeric startRow
		,required numeric startColumn
		,required numeric endRow
		,required numeric endColumn
	){
		addWorkbookArgument( arguments );
		this.getLibrary().clearCellRange( argumentCollection=arguments );
		return this;
	}

	// Ends chain - returns CellStyle object
	public any function createCellStyle( required struct format ){
		addWorkbookArgument( arguments );
		return this.getLibrary().createCellStyle( argumentCollection=arguments );
	}

	public SpreadsheetChainable function createSheet( string sheetName, overwrite=false ){
		addWorkbookArgument( arguments );
		this.getLibrary().createSheet( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function deleteColumn( required numeric column ){
		addWorkbookArgument( arguments );
		this.getLibrary().deleteColumn( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function deleteColumns( required string range ){
		addWorkbookArgument( arguments );
		this.getLibrary().deleteColumns( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function deleteRow( required numeric row ){
		addWorkbookArgument( arguments );
		this.getLibrary().deleteRow( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function deleteRows( required string range ){
		addWorkbookArgument( arguments );
		this.getLibrary().deleteRows( argumentCollection=arguments );
		return this;
	}

	// Ends chain
	public void function download( required string filename, string contentType ){
		addWorkbookArgument( arguments );
		this.getLibrary().download( argumentCollection=arguments );
	}

	public SpreadsheetChainable function formatCell(
		struct format={}
		,required numeric row
		,required numeric column
		,boolean overwriteCurrentStyle=true
		,any cellStyle
	){
		addWorkbookArgument( arguments );
		this.getLibrary().formatCell( argumentCollection=arguments );
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
		this.getLibrary().formatCellRange( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function formatColumn(
		struct format={}
		,required numeric column
		,boolean overwriteCurrentStyle=true
		,any cellStyle
	){
		addWorkbookArgument( arguments );
		this.getLibrary().formatColumn( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function formatColumns(
		struct format={}
		,required string range
		,boolean overwriteCurrentStyle=true
		,any cellStyle
	){
		addWorkbookArgument( arguments );
		this.getLibrary().formatColumns( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function formatRow(
		struct format={}
		,required numeric row
		,boolean overwriteCurrentStyle=true
		,any cellStyle
	){
		addWorkbookArgument( arguments );
		this.getLibrary().formatRow( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function formatRows(
		struct format={}
		,required string range
		,boolean overwriteCurrentStyle=true
		,any cellStyle
	){
		addWorkbookArgument( arguments );
		this.getLibrary().formatRows( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function fromCsv(){
		this.setWorkbook( this.getLibrary().workbookFromCsv( argumentCollection=arguments ) );
		return this;
	}

	public SpreadsheetChainable function fromQuery(){
		this.setWorkbook( this.getLibrary().workbookFromQuery( argumentCollection=arguments ) );
		return this;
	}

	public any function getCellComment( numeric row, numeric column ){
		// Ends chain: returns struct OR array of structs
		addWorkbookArgument( arguments );
		return this.getLibrary().getCellComment( argumentCollection=arguments );
	}

	public SpreadsheetChainable function getCellComments(){
		addWorkbookArgument( arguments );
		this.getLibrary().getCellComments( argumentCollection=arguments );
		return this;
	}

	//Ends chain
	public struct function getCellFormat( required numeric row, required numeric column ){
		addWorkbookArgument( arguments );
		return this.getLibrary().getCellFormat( argumentCollection=arguments );
	}

	public any function getCellFormula( numeric row, numeric column ){
		// Ends chain: returns string OR array of strings
		addWorkbookArgument( arguments );
		return this.getLibrary().getCellFormula( argumentCollection=arguments );
	}

	// Ends chain
	public string function getCellHyperLink( required numeric row, required numeric column ){
		addWorkbookArgument( arguments );
		return this.getLibrary().getCellHyperLink( argumentCollection=arguments );
	}

	// Ends chain
	public string function getCellType( required numeric row, required numeric column ){
		addWorkbookArgument( arguments );
		return this.getLibrary().getCellType( argumentCollection=arguments );
	}

	// Ends chain
	public any function getCellValue( required numeric row, required numeric column ){
		addWorkbookArgument( arguments );
		return this.getLibrary().getCellValue( argumentCollection=arguments );
	}

	// Ends chain
	public numeric function getColumnCount( sheetNameOrNumber ){
		addWorkbookArgument( arguments );
		return this.getLibrary().getColumnCount( argumentCollection=arguments );
	}

	// Ends chain
	public numeric function getColumnWidth( required numeric column ){
		addWorkbookArgument( arguments );
		return this.getLibrary().getColumnWidth( argumentCollection=arguments );
	}

	// Ends chain
	public numeric function getColumnWidthInPixels( required numeric column ){
		addWorkbookArgument( arguments );
		return this.getLibrary().getColumnWidthInPixels( argumentCollection=arguments );
	}

	// Ends chain
	public numeric function getLastRowNumber( sheetNameOrNumber ){
		addWorkbookArgument( arguments );
		return this.getLibrary().getLastRowNumber( argumentCollection=arguments );
	}

	// Ends chain
	public numeric function getRowCount( sheetNameOrNumber ){
		addWorkbookArgument( arguments );
		return this.getLibrary().getRowCount( argumentCollection=arguments );
	}

	public SpreadsheetChainable function hideColumn( required numeric column ){
		addWorkbookArgument( arguments );
		this.getLibrary().hideColumn( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function hideRow( required numeric row ){
		addWorkbookArgument( arguments );
		this.getLibrary().hideRow( argumentCollection=arguments );
		return this;
	}

	// Ends chain
	public struct function info(){
		// argument name is workbookOrPath not workbook so custom handling
		throwErrorIfWorkbookIsNull();
		throwErrorIfWorkbookIsInvalid();
		return this.getLibrary().info( workbookOrPath=this.getWorkbook() );
	}

	public SpreadsheetChainable function mergeCells(
		required numeric startRow
		,required numeric endRow
		,required numeric startColumn
		,required numeric endColumn
		,boolean emptyInvisibleCells=false
	){
		addWorkbookArgument( arguments );
		this.getLibrary().mergeCells( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function read(){
		this.setWorkbook( this.getLibrary().read( argumentCollection=arguments ) );
		return this;
	}

	// Ends chain
	public binary function readBinary(){
		addWorkbookArgument( arguments );
		return this.getLibrary().readBinary( argumentCollection=arguments );
	}

	public SpreadsheetChainable function removePrintGridlines(){
		addWorkbookArgument( arguments );
		this.getLibrary().removePrintGridlines( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function removeSheet( required string sheetName ){
		addWorkbookArgument( arguments );
		this.getLibrary().removeSheet( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function removeSheetNumber( required numeric sheetNumber ){
		addWorkbookArgument( arguments );
		this.getLibrary().removeSheetNumber( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function renameSheet( required string sheetName, required numeric sheetNumber ){
		addWorkbookArgument( arguments );
		this.getLibrary().renameSheet( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setActiveCell( required numeric row, required numeric column ){
		addWorkbookArgument( arguments );
		this.getLibrary().setActiveCell( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setActiveSheet( string sheetName, numeric sheetNumber ){
		addWorkbookArgument( arguments );
		this.getLibrary().setActiveSheet( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setActiveSheetNumber( numeric sheetNumber ){
		addWorkbookArgument( arguments );
		this.getLibrary().setActiveSheetNumber( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setCellComment(
		required struct comment
		,required numeric row
		,required numeric column
	){
		addWorkbookArgument( arguments );
		this.getLibrary().setCellComment( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setCellFormula(
		required string formula
		,required numeric row
		,required numeric column
	){
		addWorkbookArgument( arguments );
		this.getLibrary().setCellFormula( argumentCollection=arguments );
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
		this.getLibrary().setCellHyperlink( argumentCollection=arguments );
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
		this.getLibrary().setCellRangeValue( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setCellValue( required value, required numeric row, required numeric column, string type ){
		addWorkbookArgument( arguments );
		this.getLibrary().setCellValue( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setColumnWidth( required numeric column, required numeric width ){
		addWorkbookArgument( arguments );
		this.getLibrary().setColumnWidth( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setFitToPage( required boolean state, numeric pagesWide, numeric pagesHigh ){
		addWorkbookArgument( arguments );
		this.getLibrary().setFitToPage( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setFooter(
		string leftFooter=""
		,string centerFooter=""
		,string rightFooter=""
	){
		addWorkbookArgument( arguments );
		this.getLibrary().setFooter( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setFooterImage(
		required string position // left|center|right
		,required any image
		,string imageType
	){
		addWorkbookArgument( arguments );
		this.getLibrary().setFooterImage( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setHeader(
		string leftHeader=""
		,string centerHeader=""
		,string rightHeader=""
	){
		addWorkbookArgument( arguments );
		this.getLibrary().setHeader( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setHeaderImage(
		required string position // left|center|right
		,required any image
		,string imageType
	){
		addWorkbookArgument( arguments );
		this.getLibrary().setHeaderImage( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setReadOnly( required string password ){
		addWorkbookArgument( arguments );
		this.getLibrary().setReadOnly( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setRecalculateFormulasOnNextOpen( boolean value=true ){
		addWorkbookArgument( arguments );
		this.getLibrary().setRecalculateFormulasOnNextOpen( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setRepeatingColumns( required string columnRange ){
		addWorkbookArgument( arguments );
		this.getLibrary().setRepeatingColumns( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setRepeatingRows( required string rowRange ){
		addWorkbookArgument( arguments );
		this.getLibrary().setRepeatingRows( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setRowHeight( required numeric row, required numeric height ){
		addWorkbookArgument( arguments );
		this.getLibrary().setRowHeight( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setSheetTopMargin( required numeric marginSize, string sheetName, numeric sheetNumber ){
		addWorkbookArgument( arguments );
		this.getLibrary().setSheetTopMargin( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setSheetBottomMargin( required numeric marginSize, string sheetName, numeric sheetNumber ){
		addWorkbookArgument( arguments );
		this.getLibrary().setSheetBottomMargin( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setSheetLeftMargin( required numeric marginSize, string sheetName, numeric sheetNumber ){
		addWorkbookArgument( arguments );
		this.getLibrary().setSheetLeftMargin( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setSheetRightMargin( required numeric marginSize, string sheetName, numeric sheetNumber ){
		addWorkbookArgument( arguments );
		this.getLibrary().setSheetRightMargin( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setSheetHeaderMargin( required numeric marginSize, string sheetName, numeric sheetNumber ){
		addWorkbookArgument( arguments );
		this.getLibrary().setSheetHeaderMargin( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setSheetFooterMargin( required numeric marginSize, string sheetName, numeric sheetNumber ){
		addWorkbookArgument( arguments );
		this.getLibrary().setSheetFooterMargin( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function setSheetPrintOrientation( required string mode, string sheetName, numeric sheetNumber ){
		addWorkbookArgument( arguments );
		this.getLibrary().setSheetPrintOrientation( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function shiftColumns( required numeric start, numeric end=arguments.start, numeric offset=1 ){
		addWorkbookArgument( arguments );
		this.getLibrary().shiftColumns( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function shiftRows( required numeric start, numeric end=arguments.start, numeric offset=1 ){
		addWorkbookArgument( arguments );
		this.getLibrary().shiftRows( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function showColumn( required numeric column ){
		addWorkbookArgument( arguments );
		this.getLibrary().showColumn( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function showRow( required numeric row ){
		addWorkbookArgument( arguments );
		this.getLibrary().showRow( argumentCollection=arguments );
		return this;
	}

	public SpreadsheetChainable function write(
		required string filepath
		,boolean overwrite=false
		,string password
		,string algorithm="agile"
	){
		addWorkbookArgument( arguments );
		this.getLibrary().write( argumentCollection=arguments );
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
		this.getLibrary().writeToCsv( argumentCollection=arguments );
		return this;
	}

}