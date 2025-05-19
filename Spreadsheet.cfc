component accessors="true"{

	//"static"
	property name="version" default="5.1.0" setter="false";
	property name="osgiLibBundleVersion" default="5.4.1.0" setter="false"; //first 3 octets = POI version; increment 4th with other jar updates
	property name="osgiLibBundleSymbolicName" default="spreadsheet-cfml" setter="false";
	property name="exceptionType" default="cfsimplicity.spreadsheet" setter="false";
	//configurable
	property name="dateFormats" type="struct" setter="false";
	property name="defaultWorkbookFormat" default="binary" setter="true" getter="false";
	property name="javaLoaderDotPath" default="javaLoader.JavaLoader" setter="false";
	property name="javaLoaderName" default="" setter="false";
	property name="loadJavaClassesUsing";
	property name="returnCachedFormulaValues" type="boolean" default="true";
	property name="throwExceptionOnFormulaError" type="boolean" default="false";
	//detected state
	property name="engine" setter="false";
	property name="javaClassesLastLoadedVia" default="Nothing loaded yet";
	//Lucee osgi loader
	property name="osgiLoader" setter="false";
	// Helpers
	property name="cellHelper" setter="false";
	property name="classHelper" setter="false";
	property name="colorHelper" setter="false";
	property name="columnHelper" setter="false";
	property name="commentHelper" setter="false";
	property name="csvHelper" setter="false";
	property name="dataTypeHelper" setter="false";
	property name="dateHelper" setter="false";
	property name="exceptionHelper" setter="false";
	property name="fileHelper" setter="false";
	property name="fontHelper" setter="false";
	property name="formatHelper" setter="false";
	property name="headerImageHelper" setter="false";
	property name="hyperLinkHelper" setter="false";
	property name="imageHelper" setter="false";
	property name="infoHelper" setter="false";
	property name="queryHelper" setter="false";
	property name="rangeHelper" setter="false";
	property name="rowHelper" setter="false";
	property name="sheetHelper" setter="false";
	property name="streamingReaderHelper" setter="false";
	property name="stringHelper" setter="false";
	property name="workbookHelper" setter="false";

	public Spreadsheet function init( struct dateFormats, string javaLoaderDotPath, string loadJavaClassesUsing ){
		variables.engine = determineEngine();
		loadHelpers();
		initializeDateFormats( argumentCollection=arguments );
		initializeJavaLoading( argumentCollection=arguments );
		return this;
	}

	private void function loadHelpers(){
		variables.cellHelper = New helpers.cell( this );
		variables.classHelper = New helpers.class( this );
		variables.colorHelper = New helpers.color( this );
		variables.columnHelper = New helpers.column( this );
		variables.commentHelper = New helpers.comment( this );
		variables.csvHelper = New helpers.csv( this );
		variables.dataTypeHelper = New helpers.dataType( this );
		variables.dateHelper = New helpers.date( this );
		variables.exceptionHelper = New helpers.exception( this );
		variables.fileHelper = New helpers.file( this );
		variables.fontHelper = New helpers.font( this );
		variables.formatHelper = New helpers.format( this );
		variables.headerImageHelper = New helpers.headerImage( this );
		variables.hyperLinkHelper = New helpers.hyperLink( this );
		variables.imageHelper = New helpers.image( this );
		variables.infoHelper = New helpers.info( this );
		variables.queryHelper = New helpers.query( this );
		variables.rangeHelper = New helpers.range( this );
		variables.rowHelper = New helpers.row( this );
		variables.sheetHelper = New helpers.sheet( this );
		variables.streamingReaderHelper = New helpers.streamingReader( this );
		variables.stringHelper = New helpers.string( this );
		variables.workbookHelper = New helpers.workbook( this );
	}

	private string function determineEngine(){
		if( server.coldfusion.productname == "ColdFusion Server" )
			return "ColdFusion";
		// PLEASE NOTE: BOXLANG SUPPORT IS CURRENTLY EXPERIMENTAL AND INCOMPLETE (only circa 75% of tests pass). 
		if( server.KeyExists( "boxlang" ) )
			return "Boxlang";
		return "Lucee";
	}

	public boolean function getIsACF(){
		return variables.engine == "ColdFusion";
	}

	public boolean function getIsLucee(){
		return variables.engine == "Lucee";
	}

	public boolean function getIsBoxlang(){
		return variables.engine == "Boxlang";
	}

	public boolean function getIsACF2025(){
		return getIsACF() && ( getEngineVersion().Left( 4 ) == 2025 );
	}

	private string function getEngineVersion(){
		if( this.getIsACF() )
			return server.coldfusion.productversion;
		return server[ variables.engine.LCase() ].version;
	}

	private void function initializeDateFormats( struct dateFormats ){
		variables.dateFormats = getDateHelper().defaultFormats();
		if( arguments.KeyExists( "dateFormats" ) )
			setDateFormats( arguments.dateFormats );
	}

	private void function initializeJavaLoading(){
		//engine defaults
		if( getIsLucee() )
			variables.loadJavaClassesUsing = "osgi";
		if( getIsACF() )
			variables.loadJavaClassesUsing = "JavaLoader";
		if( getIsBoxlang() )
			variables.loadJavaClassesUsing = "dynamicPath";
		//configurable
		if( arguments.KeyExists( "loadJavaClassesUsing" ) )
			variables.loadJavaClassesUsing = getClassHelper().validateLoadingMethod( arguments.loadJavaClassesUsing );
		if( ListFindNoCase( "dynamicPath,javaSettings,classPath", variables.loadJavaClassesUsing ) )
			return;
		if( variables.loadJavaClassesUsing == "osgi" ){
			variables.osgiLoader = New osgiLoader();
			return;
		}
		variables.javaLoaderName = "spreadsheetLibraryClassLoader-#this.getVersion()#-#Hash( GetCurrentTemplatePath() )#";
		 // Option to use the dot path of an existing javaloader installation to save duplication
		if( arguments.KeyExists( "javaLoaderDotPath" ) )
			variables.javaLoaderDotPath = arguments.javaLoaderDotPath;
	}

	private string function getDefaultWorkbookFormat(){
		if( ListFindNoCase( "xml,xlsx", variables.defaultWorkbookFormat ) )
			return "xml";
		return "binary";
	}

	/* Utilities */

	public struct function getEnvironment(){
		return {
			dateFormats: this.getDateFormats()
			,engine: this.getEngine() & " " & getEngineVersion()
			,javaVersion: getJavaVersion()
			,javaLoaderDotPath: this.getJavaLoaderDotPath()
			,javaClassesLastLoadedVia: this.getJavaClassesLastLoadedVia()
			,javaLoaderName: this.getJavaLoaderName()
			,loadJavaClassesUsing: this.getLoadJavaClassesUsing()
			,version: this.getVersion()
			,poiVersion: this.getPoiVersion()
			,osgiLibBundleVersion: this.getOsgiLibBundleVersion()
		};
	}

	public string function getJavaVersion(){
		var result = [ server.system.properties[ "java.runtime.name" ]?:"" ];
		result.Append( server.system.properties[ "java.runtime.version" ]?:"" );
		return Trim( result.ToList( " " ) );
	}

	public string function getPoiVersion(){
		return createJavaObject( "org.apache.poi.Version" ).getVersion();
	}

	public string function getLibPath(){
		return GetDirectoryFromPath( GetCurrentTemplatePath() ) & "lib/";
	}

	public JavaLoader function getJavaLoaderInstance(){
		/* Not in classHelper because of difficulty of accessing JL via dot path from there */
		if( server.KeyExists( this.getJavaLoaderName() ) )
			return server[ this.getJavaLoaderName() ];
		server[ this.getJavaLoaderName() ] = CreateObject( "component", this.getJavaLoaderDotPath() ).init( loadPaths=DirectoryList( getLibPath() ), loadColdFusionClassPath=false, trustedSource=true );
		return server[ this.getJavaLoaderName() ];
	}

	public Spreadsheet function flushPoiLoader(){
		lock scope="server" timeout="10" {
			StructDelete( server, this.getJavaLoaderName() );
		};
		return this;
	}

	public Spreadsheet function flushOsgiBundle( string version ){
		if( variables.loadJavaClassesUsing != "osgi" )
			return this;
		var allBundles = getOsgiLoader().getCFMLEngineFactory().getBundleContext().getBundles();
		var spreadsheetBundles = ArrayFilter( allBundles, function( bundle ){
			return ( bundle.getSymbolicName() == this.getOsgiLibBundleSymbolicName() );
		});
		if( arguments.KeyExists( "version" ) ){
			getOsgiLoader().uninstallBundle( this.getOsgiLibBundleSymbolicName(), arguments.version );
			return this;
		}
		for( var bundle in spreadsheetBundles ){
			getOsgiLoader().uninstallBundle( this.getOsgiLibBundleSymbolicName(), bundle.getVersion() );
		}
		return this;
	}

	public any function createJavaObject( required string className ){
		return getClassHelper().loadClass( arguments.className );
	}

	/* check physical path of a specific class */
	public void function dumpPathToClass( required string className ){
		if( IsNull( getOsgiLoader() ) )
			return getClassHelper().dumpPathToClassNoOsgi( arguments.className );
		var bundle = getOsgiLoader().getBundle( this.getOsgiLibBundleSymbolicName(), this.getOsgiLibBundleVersion() );
		var poi = createJavaObject( "org.apache.poi.Version" );
		var path = BundleInfo( poi ).location & "!" &  bundle.getResource( arguments.className.Replace( ".", "/", "all" ) & ".class" ).getPath();
		WriteDump( path );
	}

	public numeric function getWorkbookCellStylesTotal( required workbook ){
		return arguments.workbook.getNumCellStyles(); // limit is 4K xls/64K xlsx
	}

	public Spreadsheet function clearCellStyleCache(){
		getFormatHelper().initCellStyleCache();
		return this;
	}

	public struct function getCellStyleCache(){
		return getFormatHelper().getCachedCellStyles();
	}

	/* MAIN PUBLIC API */

	public Spreadsheet function addAutofilter( required workbook, string cellRange="", numeric row=1 ){
		arguments.cellRange = arguments.cellRange.Trim();
		if( arguments.cellRange.IsEmpty() ){
			//default to all columns in the first (default) or specified row
			var rowIndex = ( Max( 0, arguments.row -1 ) );
			var cellRangeAddress = getRangeHelper().getCellRangeAddressFromRowIndex( arguments.workbook, rowIndex );
			getSheetHelper().getActiveSheet( arguments.workbook ).setAutoFilter( cellRangeAddress );
			return this;
		}
		getSheetHelper().getActiveSheet( arguments.workbook ).setAutoFilter( getRangeHelper().getCellRangeAddressFromReference( arguments.cellRange ) );
		return this;
	}

	public Spreadsheet function addColumn(
		required workbook
		,required data // Delimited list of values OR array
		,numeric startRow
		,numeric startColumn
		,boolean insert=false
		,string delimiter=","
		,boolean autoSize=false
		,string datatype
	){
		var sheet = getSheetHelper().getActiveSheet( arguments.workbook );
		var rowIndex = arguments.KeyExists( "startRow" )? ( arguments.startRow -1 ): 0;
		var columnIndex = getColumnHelper().getNewColumnIndex( sheet, rowIndex, arguments.startColumn?:0 );
		if( arguments.autoSize )
			var columnNumber = ( columnIndex +1 ); //stash the starting column number
		var columnData = IsArray( arguments.data )? arguments.data: ListToArray( arguments.data, arguments.delimiter );//Don't use ListToArray() member function: value may not support it
		for( var cellValue in columnData ){
			var row = sheet.getRow( rowIndex );
			if( rowIndex > getSheetHelper().getLastRowIndex( sheet ) || IsNull( row ) )
				row = getRowHelper().createRow( arguments.workbook, rowIndex );
			// NB: row.getLastCellNum() returns the cell index PLUS ONE or -1 if not found
			var insertRequired = ( arguments.KeyExists( "startColumn" ) && arguments.insert && ( columnIndex < row.getLastCellNum() ) );
			if( insertRequired )
				getColumnHelper().shiftColumnsRightStartingAt( columnIndex, row, arguments.workbook );
			var cellValueArgs = {
				workbook: arguments.workbook
				,cell: getCellHelper().createCell( row, columnIndex )
				,value: cellValue
			};
			if( arguments.KeyExists( "datatype" ) )
				cellValueArgs.type = arguments.datatype;
			getCellHelper().setCellValueAsType( argumentCollection=cellValueArgs );
			rowIndex++;
		}
		if( arguments.autoSize )
			autoSizeColumn( arguments.workbook, columnNumber );
		return this;
	}

	public Spreadsheet function addConditionalFormatting( required workbook, required ConditionalFormatting conditionalFormatting ){
		arguments.conditionalFormatting.addToWorkbook( arguments.workbook );
		return this;
	}

	public Spreadsheet function addDataValidation( required workbook, required DataValidation dataValidation ){
		arguments.dataValidation.addToWorkbook( arguments.workbook );
		return this;
	}

	public Spreadsheet function addFreezePane(
		required workbook
		,required numeric freezeColumn
		,required numeric freezeRow
		,numeric leftmostColumn //left column visible in right pane
		,numeric topRow //top row visible in bottom pane
	){
		var sheet = getSheetHelper().getActiveSheet( arguments.workbook );
		if( arguments.KeyExists( "leftmostColumn" ) && !arguments.KeyExists( "topRow" ) )
			arguments.topRow = arguments.freezeRow;
		if( arguments.KeyExists( "topRow" ) && !arguments.KeyExists( "leftmostColumn" ) )
			arguments.leftmostColumn = arguments.freezeColumn;
		/* createFreezePane() operates on the logical row/column numbers as opposed to physical, so no need for n-1 stuff here */
		if( !arguments.KeyExists( "leftmostColumn" ) ){
			sheet.createFreezePane( JavaCast( "int", arguments.freezeColumn ), JavaCast( "int", arguments.freezeRow ) );
			return this;
		}
		sheet.createFreezePane(
			JavaCast( "int", arguments.freezeColumn )
			,JavaCast( "int", arguments.freezeRow )
			,JavaCast( "int", arguments.leftmostColumn )
			,JavaCast( "int", arguments.topRow )
		);
		return this;
	}

	public Spreadsheet function addImage(
		required workbook
		,string filepath
		,imageData
		,string imageType
		,required string anchor
	){
		var anchorCoordinates = arguments.anchor.ListToArray();
		var numberOfAnchorCoordinates = anchorCoordinates.Len();
		if( ( numberOfAnchorCoordinates != 4 ) && ( numberOfAnchorCoordinates != 8 ) )
			Throw( type=this.getExceptionType() & ".invalidAnchorArgument", message="Invalid anchor argument", detail="The anchor argument must be a comma-delimited list of integers with either 4 or 8 elements" );
		var args = { workbook: arguments.workbook };
		if( arguments.KeyExists( "image" ) )
			args.image = arguments.image;//new alias instead of filepath/imageData
		if( arguments.KeyExists( "filepath" ) )
			args.image = arguments.filepath;
		if( arguments.KeyExists( "imageData" ) )
			args.image = arguments.imageData;
		if( arguments.KeyExists( "imageType" ) )
			args.imageType = arguments.imageType;
		if( !args.KeyExists( "image" ) )
			Throw( type=this.getExceptionType() & ".missingImageArgument", message="Missing image path or object", detail="Please supply either the 'filepath' or 'imageData' argument" );
		var imageIndex = getImageHelper().addImageToWorkbook( argumentCollection=args );
		var anchorObject = getImageHelper().createAnchor( arguments.workbook, anchorCoordinates );
		/* (legacy note from spreadsheet extension) TODO: need to look into createDrawingPatriarch() vs. getDrawingPatriarch() since create will kill any existing images. getDrawingPatriarch() throws  a null pointer exception when an attempt is made to add a second image to the spreadsheet  */
		getSheetHelper().getActiveSheet( arguments.workbook ).createDrawingPatriarch().createPicture( anchorObject, imageIndex );
		return this;
	}

	public Spreadsheet function addInfo( required workbook, required struct info ){
		// Valid struct keys are author, category, lastauthor, comments, keywords, manager, company, subject, title
		if( isBinaryFormat( arguments.workbook ) ){
			getInfoHelper().addInfoBinary( arguments.workbook, arguments.info );
			return this;
		}
		getInfoHelper().addInfoXml( arguments.workbook, arguments.info );
		return this;
	}

	public Spreadsheet function addPageBreaks( required workbook, string rowBreaks="", string columnBreaks="" ){
		arguments.rowBreaks = Trim( arguments.rowBreaks ); //Don't use member function in case value is in fact numeric
		arguments.columnBreaks = Trim( arguments.columnBreaks );
		if( arguments.rowBreaks.IsEmpty() && arguments.columnBreaks.IsEmpty() )
			Throw( type=this.getExceptionType() & ".missingRowOrColumnBreaksArgument", message="Missing row or column breaks argument", detail="You must specify the rows and/or columns at which page breaks should be added." );
		arguments.rowBreaks = arguments.rowBreaks.ListToArray();
		arguments.columnBreaks = arguments.columnBreaks.ListToArray();
		var sheet = getSheetHelper().getActiveSheet( arguments.workbook );
		sheet.setAutoBreaks( false ); // Not sure if this is necessary: https://stackoverflow.com/a/14900320/204620
		for( var rowNumber in arguments.rowBreaks )
			sheet.setRowBreak( JavaCast( "int", ( rowNumber -1 ) ) );
		for( var columnNumber in arguments.columnBreaks )
			sheet.setColumnBreak( JavaCast( "int", ( columnNumber -1 ) ) );
		return this;
	}

	public Spreadsheet function addPrintGridlines( required workbook ){
		getSheetHelper().getActiveSheet( arguments.workbook ).setPrintGridlines( JavaCast( "boolean", true ) );
		return this;
	}

	public Spreadsheet function addRow(
		required workbook
		,required data // Delimited list of data, OR array
		,numeric row
		,numeric column=1
		,boolean insert=true
		,string delimiter=","
		,boolean handleEmbeddedCommas=true // When true, values enclosed in single quotes are treated as a single element like in ACF. Only applies when the delimiter is a comma.
		,boolean autoSizeColumns=false
		,struct datatypes
	){
		if( !IsArray( arguments.data ) )
			arguments.data = getRowHelper().parseListDataToArray( arguments.data, arguments.delimiter, arguments.handleEmbeddedCommas );
		arguments.data = [ arguments.data ];// array of arrays for addRows()
		return addRows( argumentCollection=arguments );
	}

	public Spreadsheet function addRows(
		required workbook
		,required data // query or array of arrays
		,numeric row
		,numeric column=1
		,boolean insert=true
		,boolean autoSizeColumns=false
		,boolean includeQueryColumnNames=false
		,boolean ignoreQueryColumnDataTypes=false
		,struct datatypes
	){
		if( arguments.KeyExists( "row" ) && ( arguments.row <= 0 ) )
			Throw( type=this.getExceptionType() & ".invalidRowArgument", message="Invalid row value", detail="The value for row must be greater than or equal to 1." );
		if( arguments.KeyExists( "column" ) && ( arguments.column <= 0 ) )
			Throw( type=this.getExceptionType() & ".invalidColumnArgument", message="Invalid column value", detail="The value for column must be greater than or equal to 1." );
		if( !arguments.insert && !arguments.KeyExists( "row") )
			Throw( type=this.getExceptionType() & ".missingRowArgument", message="Missing row value", detail="To replace a row using 'insert', please specify the row to replace." );
		var dataIsQuery = IsQuery( arguments.data );
		var dataIsArray = IsArray( arguments.data );
		if( !dataIsQuery && !dataIsArray )
			Throw( type=this.getExceptionType() & ".invalidDataArgument", message="Invalid data argument", detail="The data passed in must be either a query or an array of row arrays." );
		getDataTypeHelper().checkDataTypesArgument( arguments );
		var totalRows = dataIsQuery? arguments.data.recordCount: arguments.data.Len();
		if( totalRows == 0 )
			return this;
		// array data must be an array of arrays, not structs
		if( dataIsArray && !IsArray( arguments.data[ 1 ] ) )
			Throw( type=this.getExceptionType() & ".invalidDataArgument", message="Invalid data argument", detail="Data passed as an array must be an array of arrays, one per row" );
		var sheet = getSheetHelper().getActiveSheet( arguments.workbook );
		var nextRowIndex = getSheetHelper().getNextEmptyRowIndex( sheet );
		var insertAtRowIndex = arguments.KeyExists( "row" )? arguments.row -1: nextRowIndex;
		if( arguments.KeyExists( "row" ) && ( arguments.row <= nextRowIndex ) && arguments.insert )
			shiftRows( arguments.workbook, arguments.row, nextRowIndex, totalRows );
		var currentRowIndex = insertAtRowIndex;
		var overrideDataTypes = arguments.KeyExists( "datatypes" );
		if( arguments.autoSizeColumns && isStreamingXmlFormat( arguments.workbook ) )
			getSheetHelper().getActiveSheet( arguments.workbook ).trackAllColumnsForAutoSizing();
			/* this will affect performance but is needed for autoSizeColumns to work properly with SXSSF: https://poi.apache.org/apidocs/dev/org/apache/poi/xssf/streaming/SXSSFSheet.html#trackAllColumnsForAutoSizing */
		var addRowsArgs = {
			workbook: arguments.workbook
			,data: arguments.data
			,column: arguments.column
			,includeQueryColumnNames: arguments.includeQueryColumnNames
			,ignoreQueryColumnDataTypes: arguments.ignoreQueryColumnDataTypes
			,autoSizeColumns: arguments.autoSizeColumns
			,currentRowIndex: currentRowIndex
			,overrideDataTypes: overrideDataTypes
		};
		if( overrideDataTypes )
			addRowsArgs.datatypes = arguments.datatypes;
		if( dataIsQuery ){
			getRowHelper().addRowsFromQuery( argumentCollection=addRowsArgs );
			return this;
		}
		getRowHelper().addRowsFromArray( argumentCollection=addRowsArgs );
   	return this;
	}

	public Spreadsheet function addSplitPane(
		required workbook
		,required numeric xSplitPosition
		,required numeric ySplitPosition
		,required numeric leftmostColumn
		,required numeric topRow
		,string activePane="UPPER_LEFT" //Valid values are LOWER_LEFT, LOWER_RIGHT, UPPER_LEFT, and UPPER_RIGHT
	){
		var sheet = getSheetHelper().getActiveSheet( arguments.workbook );
		arguments.activePane = sheet[ "PANE_#arguments.activePane#" ];
		sheet.createSplitPane(
			JavaCast( "int", arguments.xSplitPosition )
			,JavaCast( "int", arguments.ySplitPosition )
			,JavaCast( "int", arguments.leftmostColumn )
			,JavaCast( "int", arguments.topRow )
			,JavaCast( "int", arguments.activePane )
		);
		return this;
	}

	public Spreadsheet function autoSizeColumn( required workbook, required numeric column, boolean useMergedCells=false ){
		if( arguments.column <= 0 )
			Throw( type=this.getExceptionType() & ".invalidColumnArgument", message="Invalid column argument", detail="The value for column must be greater than or equal to 1." );
		// Adjusts the width of the specified column to fit the contents. For performance reasons, this should normally be called only once per column.
		var columnIndex = ( arguments.column -1 );
		if( isStreamingXmlFormat( arguments.workbook ) )
			getSheetHelper().getActiveSheet( arguments.workbook ).trackColumnForAutoSizing( JavaCast( "int", columnIndex ) );
			// has no effect if tracking is already on
		getSheetHelper().getActiveSheet( arguments.workbook ).autoSizeColumn( columnIndex, arguments.useMergedCells );
		return this;
	}

	public any function binaryFromQuery(
		required query data
		,boolean addHeaderRow=true
		,boolean boldHeaderRow=true
		,boolean xmlFormat=false
		,boolean streamingXml=false
		,numeric streamingWindowSize=100
		,boolean ignoreQueryColumnDataTypes=false
		,struct datatypes
	){
		var workbook = workbookFromQuery( argumentCollection=arguments );
		var binary = readBinary( workbook );
		cleanUpStreamingXml( workbook );
		return binary;
	}

	public Spreadsheet function cleanUpStreamingXml( required workbook ){
		// SXSSF uses temporary files which MUST be cleaned up, see http://poi.apache.org/components/spreadsheet/how-to.html#sxssf
		if( isStreamingXmlFormat( arguments.workbook ) )
			arguments.workbook.dispose();
		return this;
	}

	public Spreadsheet function clearCell( required workbook, required numeric row, required numeric column ){
		// Clears the specified cell of all styles and values
		var rowObject = getRowHelper().getRowFromActiveSheet( arguments.workbook, arguments.row );
		if( IsNull( rowObject ) )
			return this;
		var columnIndex = ( arguments.column -1 );
		var cell = rowObject.getCell( JavaCast( "int", columnIndex ) );
		if( IsNull( cell ) )
			return this;
		var defaultStyle = arguments.workbook.getCellStyleAt( JavaCast( "short", 0 ) );
		cell.setCellStyle( defaultStyle );
		cell.setBlank();
		return this;
	}

	public Spreadsheet function clearCellRange(
		required workbook
		,required numeric startRow
		,required numeric startColumn
		,required numeric endRow
		,required numeric endColumn
	){
		for( var rowNumber = arguments.startRow; rowNumber <= arguments.endRow; rowNumber++ ){
			for( var columnNumber = arguments.startColumn; columnNumber <= arguments.endColumn; columnNumber++ ){
				clearCell( arguments.workbook, rowNumber, columnNumber );
			}
		}
		return this;
	}

	public SpreadSheet function collapseColumnGroup( required workbook, required numeric column ){
		getSheetHelper().getActiveSheet( arguments.workbook ).setColumnGroupCollapsed( JavaCast( "int", arguments.column-1 ), JavaCast( "boolean", true ) );
		return this;
	}

	public SpreadSheet function collapseRowGroup( required workbook, required numeric row ){
		getSheetHelper().getActiveSheet( arguments.workbook ).setRowGroupCollapsed( JavaCast( "int", arguments.row-1 ), JavaCast( "boolean", true ) );
		return this;
	}

	public any function createCellStyle( required workbook, required struct format ){
		return getFormatHelper().buildCellStyle( arguments.workbook, arguments.format );
	}

	public Spreadsheet function createSheet( required workbook, string sheetName, overwrite=false ){
		local.sheetName = getSheetHelper().createOrValidateSheetName( argumentCollection=arguments );
		if( !getSheetHelper().sheetExists( workbook=arguments.workbook, sheetName=sheetName ) ){
			arguments.workbook.createSheet( JavaCast( "String", sheetName ) );
			return this;
		}
		// sheet already exists with that name
		if( !arguments.overwrite )
			Throw( type=this.getExceptionType() & ".sheetNameAlreadyExists", message="Sheet name already exists", detail="A sheet with the name '#sheetName#' already exists in this workbook" );
		// OK to replace the existing
		var sheetIndexToReplace = arguments.workbook.getSheetIndex( JavaCast( "string", sheetName ) );
		getSheetHelper().deleteSheetAtIndex( arguments.workbook, sheetIndexToReplace );
		var newSheet = arguments.workbook.createSheet( JavaCast( "String", sheetName ) );
		var moveToIndex = sheetIndexToReplace;
		getSheetHelper().moveSheet( arguments.workbook, sheetName, moveToIndex );
		return this;
	}

	public query function csvToQuery(
		string csv=""
		,string filepath=""
		,boolean firstRowIsHeader=false
		,boolean trim=true
		,string delimiter
		,array queryColumnNames
		,any queryColumnTypes="" //'auto', single default type e.g. 'VARCHAR', or list of types, or struct of column names/types mapping. Empty means no types are specified.
		,boolean makeColumnNamesSafe=false
	){
		var csvIsString = arguments.csv.Len();
		var csvIsFile = arguments.filepath.Len();
		if( !csvIsString && !csvIsFile )
			Throw( type=this.getExceptionType() & ".missingRequiredArgument", message="Missing required argument", detail="Please provide either a csv string (csv), or the path of a file containing one (filepath)." );
		if( csvIsString && csvIsFile )
			Throw( type=this.getExceptionType() & ".invalidArgumentCombination", message="Mutually exclusive arguments: 'csv' and 'filepath'", detail="Only one of either 'filepath' or 'csv' arguments may be provided." );
		if( IsStruct( arguments.queryColumnTypes ) && !arguments.firstRowIsHeader && !arguments.KeyExists( "queryColumnNames" )  )
			Throw( type=this.getExceptionType() & ".invalidArgumentCombination", message="Invalid argument 'queryColumnTypes'.", detail="When specifying 'queryColumnTypes' as a struct you must also set the 'firstRowIsHeader' argument to true OR provide 'queryColumnNames'" );
		var format = getCsvHelper().getFormat( arguments.delimiter?:"" );
		var parsed = csvIsFile?
			getCsvHelper().parseFromFile( arguments.filepath, arguments.trim, format ):
			getCsvHelper().parseFromString( arguments.csv, arguments.trim, format );
		var data = parsed.data;
		var maxColumnCount = parsed.maxColumnCount;
		if( arguments.KeyExists( "queryColumnNames" ) && arguments.queryColumnNames.Len() ){
			var columnNames = arguments.queryColumnNames;
			var parsedQueryColumnTypes = getQueryHelper().parseQueryColumnTypesArgument( arguments.queryColumnTypes, columnNames, maxColumnCount, data );
			return getQueryHelper()._QueryNew( columnNames, parsedQueryColumnTypes, data, arguments.makeColumnNamesSafe );
		}
		var columnNames = getCsvHelper().getColumnNames( arguments.firstRowIsHeader, data, maxColumnCount );
		if( arguments.firstRowIsHeader )
			data.DeleteAt( 1 );
		var parsedQueryColumnTypes = getQueryHelper().parseQueryColumnTypesArgument( arguments.queryColumnTypes, columnNames, maxColumnCount, data );
		return getQueryHelper()._QueryNew( columnNames, parsedQueryColumnTypes, data, arguments.makeColumnNamesSafe );
	}

	public Spreadsheet function deleteColumn( required workbook, required numeric column ){
		if( arguments.column <= 0 )
			Throw( type=this.getExceptionType() & ".invalidColumnArgument", message="Invalid column argument", detail="The value for column must be greater than or equal to 1." );
			// POI doesn't have remove column functionality, so iterate over all the rows and remove the column indicated
		var rowIterator = getSheetHelper().getActiveSheetRowIterator( arguments.workbook );
		var columnIndex = ( arguments.column -1 );
		while( rowIterator.hasNext() ){
			var row = rowIterator.next();
			var cell = row.getCell( JavaCast( "int", columnIndex ) );
			if( IsNull( cell ) )
				continue;
			row.removeCell( cell );
		}
		return this;
	}

	public Spreadsheet function deleteColumns( required workbook, required string range ){
		// Validate and extract the ranges. Range is a comma-delimited list of ranges, and each value can be either a single number or a range of numbers with a hyphen.
		var allRanges = getRangeHelper().extractRanges( arguments.range, arguments.workbook, "column" );
		for( var thisRange in allRanges ){
			if( thisRange.startAt == thisRange.endAt ){ // Just one row
				deleteColumn( arguments.workbook, thisRange.startAt );
				continue;
			}
			for( var columnNumber = thisRange.startAt; columnNumber <= thisRange.endAt; columnNumber++ )
				deleteColumn( arguments.workbook, columnNumber );
		}
		return this;
	}

	public Spreadsheet function deleteRow( required workbook, required numeric row ){
		// Deletes the data from a row. Does not physically delete the row
		if( arguments.row <= 0 )
			Throw( type=this.getExceptionType() & ".invalidRowArgument", message="Invalid row argument", detail="The value for row must be greater than or equal to 1." );
		var sheet = getSheetHelper().getActiveSheet( arguments.workbook );
		var rowIndex = ( arguments.row -1 );
		if( !getRowHelper().rowExists( rowIndex, sheet ) )
			return this;
		sheet.removeRow( sheet.getRow( rowIndex ) );
		return this;
	}

	public Spreadsheet function deleteRows( required workbook, required string range ){
		// Validate and extract the ranges. Range is a comma-delimited list of ranges, and each value can be either a single number or a range of numbers with a hyphen.
		var allRanges = getRangeHelper().extractRanges( arguments.range, arguments.workbook );
		for( var thisRange in allRanges ){
			if( thisRange.startAt == thisRange.endAt ){ // Just one row
				deleteRow( arguments.workbook, thisRange.startAt );
				continue;
			}
			for( var rowNumber = thisRange.startAt; rowNumber <= thisRange.endAt; rowNumber++ )
				deleteRow( arguments.workbook, rowNumber );
		}
		return this;
	}

	public void function download( required workbook, required string filename, string contentType ){
		var safeFilename = getFileHelper().filenameSafe( arguments.filename );
		var filenameWithoutExtension = safeFilename.REReplace( "\.xlsx?$", "" );
		var extension = isXmlFormat( arguments.workbook )? "xlsx": "xls";
		arguments.filename = filenameWithoutExtension & "." & extension;
		var binary = readBinary( arguments.workbook );
		cleanUpStreamingXml( arguments.workbook );
		if( !arguments.KeyExists( "contentType" ) )
			arguments.contentType = isXmlFormat( arguments.workbook )? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "application/msexcel";
		getFileHelper().downloadBinaryVariable( binary, arguments.filename, arguments.contentType );
	}

	public void function downloadCsvFromFile(
		required string src
		,required string filename
		,string contentType="text/csv"
		,string columns
		,string columnNames
		,numeric headerRow
		,string rows
		,string sheetName
		,numeric sheetNumber // 1-based
		,boolean includeHeaderRow=false
		,boolean includeBlankRows=false
		,boolean fillMergedCellsWithVisibleValue=false
		,string delimiter=","
	){
		arguments.format = "csv";
		arguments.csvDelimiter = arguments.delimiter;
		var csv = read( argumentCollection=arguments );
		var binary = ToBinary( ToBase64( csv.Trim() ) );
		var safeFilename = getFileHelper().filenameSafe( arguments.filename );
		var filenameWithoutExtension = safeFilename.REReplace( "\.csv$","" );
		var extension = "csv";
		arguments.filename = filenameWithoutExtension & "." & extension;
		getFileHelper().downloadBinaryVariable( binary, arguments.filename, arguments.contentType );
	}

	public void function downloadFileFromQuery(
		required query data
		,required string filename
		,boolean addHeaderRow=true
		,boolean boldHeaderRow=true
		,boolean xmlFormat=false
		,string contentType
		,boolean streamingXml=false
		,numeric streamingWindowSize=100
		,boolean ignoreQueryColumnDataTypes=false
		,struct datatypes
	){
		var safeFilename = getFileHelper().filenameSafe( arguments.filename );
		var filenameWithoutExtension = safeFilename.REReplace( "\.xlsx?$","" );
		var extension = ( arguments.xmlFormat || arguments.streamingXml )? "xlsx": "xls";
		arguments.filename = filenameWithoutExtension & "." & extension;
		var binaryFromQueryArgs = {
			data: arguments.data
			,addHeaderRow: arguments.addHeaderRow
			,boldHeaderRow: arguments.boldHeaderRow
			,xmlFormat: arguments.xmlFormat
			,streamingXml: arguments.streamingXml
			,streamingWindowSize: arguments.streamingWindowSize
			,ignoreQueryColumnDataTypes: arguments.ignoreQueryColumnDataTypes
		};
		if( arguments.KeyExists( "datatypes" ) )
			binaryFromQueryArgs.datatypes = arguments.datatypes;
		var binary = binaryFromQuery( argumentCollection=binaryFromQueryArgs );
		if( !arguments.KeyExists( "contentType" ) )
			arguments.contentType = arguments.xmlFormat? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "application/msexcel";
		getFileHelper().downloadBinaryVariable( binary, arguments.filename, arguments.contentType );
	}

	public SpreadSheet function expandColumnGroup( required workbook, required numeric column ){
		getSheetHelper().getActiveSheet( arguments.workbook ).setColumnGroupCollapsed( JavaCast( "int", arguments.column-1 ), JavaCast( "boolean", false ) );
		return this;
	}

	public SpreadSheet function expandRowGroup( required workbook, required numeric row ){
		/*
			20250413: bug in POI XSSF? Expansion fails unless row is unhidden first. Also errors if group begins on row 1
		*/
		if( isXmlFormat( arguments.workbook ) && isRowHidden( arguments.workbook, arguments.row ) )
			showRow( arguments.workbook, arguments.row );
		getSheetHelper().getActiveSheet( arguments.workbook ).setRowGroupCollapsed( JavaCast( "int", arguments.row-1 ), JavaCast( "boolean", false ) );
		return this;
	}

	public Spreadsheet function formatCell(
		required workbook
		,any format //struct or cellStyle
		,required numeric row
		,required numeric column
		,boolean overwriteCurrentStyle=true
	){
		arguments = getFormatHelper().checkFormatArguments( argumentCollection=arguments );
		var cell = getCellHelper().initializeCell( arguments.workbook, arguments.row, arguments.column );
		if( arguments.KeyExists( "cellStyle" ) ){
			getFormatHelper().setCellStyle( cell, arguments.cellStyle );
			return this;
		}
		var cellStyle = arguments.overwriteCurrentStyle?
			getFormatHelper().getCachedCellStyle( arguments.workbook, arguments.format ):
			getFormatHelper().buildCellStyle( arguments.workbook, arguments.format, cell.getCellStyle() );
		getFormatHelper().setCellStyle( cell, cellStyle, arguments.format );
		return this;
	}

	public Spreadsheet function formatCellRange(
		required workbook
		,any format //struct or cellStyle
		,required numeric startRow
		,required numeric endRow
		,required numeric startColumn
		,required numeric endColumn
		,boolean overwriteCurrentStyle=true
	){
		arguments = getFormatHelper().checkFormatArguments( argumentCollection=arguments );
		var formatCellArgs = {
			workbook: arguments.workbook
			,format: arguments.format?:arguments?.cellStyle
			,overwriteCurrentStyle: arguments.overwriteCurrentStyle
		};
		for( var rowNumber = arguments.startRow; rowNumber <= arguments.endRow; rowNumber++ ){
			for( var columnNumber = arguments.startColumn; columnNumber <= arguments.endColumn; columnNumber++ )
				formatCell( argumentCollection=formatCellArgs, row=rowNumber, column=columnNumber );
		}
		return this;
	}

	public Spreadsheet function formatColumn(
		required workbook
		,any format //struct or cellStyle
		,required numeric column
		,boolean overwriteCurrentStyle=true
	){
		arguments = getFormatHelper().checkFormatArguments( argumentCollection=arguments );
		if( arguments.column < 1 )
			Throw( type=this.getExceptionType() & ".invalidColumnArgument", message="Invalid column argument", detail="The column value must be greater than 0" );
		var formatCellArgs = {
			workbook: arguments.workbook
			,format: arguments.format?:arguments?.cellStyle
			,column: arguments.column
			,overwriteCurrentStyle: arguments.overwriteCurrentStyle
		};
		var rowIterator = getSheetHelper().getActiveSheetRowIterator( arguments.workbook );
		while( rowIterator.hasNext() ){
			var rowNumber = rowIterator.next().getRowNum() + 1;
			formatCell( argumentCollection=formatCellArgs, row=rowNumber );
		}
		return this;
	}

	public Spreadsheet function formatColumns(
		required workbook
		,any format //struct or cellStyle
		,required string range
		,boolean overwriteCurrentStyle=true
	){
		arguments = getFormatHelper().checkFormatArguments( argumentCollection=arguments );
		// Validate and extract the ranges. Range is a comma-delimited list of ranges, and each value can be either a single number or a range of numbers with a hyphen.
		var allRanges = getRangeHelper().extractRanges( arguments.range, arguments.workbook, "column" );
		var formatColumnArgs = {
			workbook: arguments.workbook
			,format: arguments.format?:arguments?.cellStyle
			,overwriteCurrentStyle: arguments.overwriteCurrentStyle
		};
		for( var thisRange in allRanges ){
			for( var columnNumber = thisRange.startAt; columnNumber <= thisRange.endAt; columnNumber++ ){
				formatColumn( argumentCollection=formatColumnArgs, column=columnNumber );
			}
		}
		return this;
	}

	public Spreadsheet function formatRow(
		required workbook
		,any format //struct or cellStyle
		,required numeric row
		,boolean overwriteCurrentStyle=true
	){
		arguments = getFormatHelper().checkFormatArguments( argumentCollection=arguments );
		var theRow = getRowHelper().getRowFromActiveSheet( arguments.workbook, arguments.row );
		if( IsNull( theRow ) )
			return this;
		var formatCellArgs = {
			workbook: arguments.workbook
			,format: arguments.format?:arguments?.cellStyle
			,row: arguments.row
			,overwriteCurrentStyle: arguments.overwriteCurrentStyle
		};
		var cellIterator = theRow.cellIterator();
		while( cellIterator.hasNext() ){
			var columnNumber = ( cellIterator.next().getColumnIndex() +1 );
			formatCell( argumentCollection=formatCellArgs, column=columnNumber );
		}
		return this;
	}

	public Spreadsheet function formatRows(
		required workbook
		,any format //struct or cellStyle
		,required string range
		,boolean overwriteCurrentStyle=true
	){
		arguments = getFormatHelper().checkFormatArguments( argumentCollection=arguments );
		// Validate and extract the ranges. Range is a comma-delimited list of ranges, and each value can be either a single number or a range of numbers with a hyphen.
		var allRanges = getRangeHelper().extractRanges( arguments.range, arguments.workbook );
		var formatRowArgs = {
			workbook: arguments.workbook
			,format: arguments.format?:arguments?.cellStyle
			,overwriteCurrentStyle: arguments.overwriteCurrentStyle
		};
		for( var thisRange in allRanges ){
			for( var rowNumber = thisRange.startAt; rowNumber <= thisRange.endAt; rowNumber++ ){
				formatRow( argumentCollection=formatRowArgs, row=rowNumber );
			}
		}
		return this;
	}

	public struct function getActiveCell( required workbook ){
		var activeCell = getSheetHelper().getActiveSheet( arguments.workbook ).getActiveCell();
		return { column: activeCell.getColumn()+1, row: activeCell.getRow()+1 };
	}

	public string function getCellAddress( required workbook, required numeric row, required numeric column ){
		var cell = getCellHelper().getCellAt( arguments.workbook, arguments.row, arguments.column );
		if( IsNull( cell ) )
			getExceptionHelper().throwInvalidCellException( arguments.row, arguments.column );
		return cell.getAddress().formatAsString();
	}

	public any function getCellComment( required workbook, numeric row, numeric column ){
		// returns struct OR array of structs
		if( arguments.KeyExists( "row" ) && !arguments.KeyExists( "column" ) )
			Throw( type=this.getExceptionType() & ".invalidArgumentCombination", message="Invalid argument combination", detail="If you specify the row you must also specify the column" );
		if( arguments.KeyExists( "column" ) && !arguments.KeyExists( "row" ) )
			Throw( type=this.getExceptionType() & ".invalidArgumentCombination", message="Invalid argument combination", detail="If you specify the column you must also specify the row" );
		if( !arguments.KeyExists( "row" ) )
			return getCellComments( arguments.workbook );// row and column weren't provided so return all the comments as an array of structs
		var cell = getCellHelper().getCellAt( arguments.workbook, arguments.row, arguments.column );
		if( IsNull( cell ) )
			getExceptionHelper().throwInvalidCellException( arguments.row, arguments.column );
		var commentObject = cell.getCellComment();
		if( IsNull( commentObject ) )
			return {};
		return {
			author: commentObject.getAuthor()
			,comment: commentObject.getString().getString()
			,column: arguments.column
			,row: arguments.row
		};
	}

	public array function getCellComments( required workbook ){
		var comments = [];
		var commentsIterator = getSheetHelper().getActiveSheet( arguments.workbook ).getCellComments().values().iterator();
		while( commentsIterator.hasNext() ){
			var commentObject = commentsIterator.next();
			var comment = {
				author: commentObject.getAuthor()
				,comment: commentObject.getString().getString()
				,column: ( commentObject.getColumn() +1 )
				,row: ( commentObject.getRow() +1 )
			};
			comments.Append( comment );
		}
		return comments;
	}

	public struct function getCellFormat( required workbook, required numeric row, required numeric column ){
		var cell = getCellHelper().getCellAt( arguments.workbook, arguments.row, arguments.column );
		if( IsNull( cell ) )
			getExceptionHelper().throwInvalidCellException( arguments.row, arguments.column );
		var cellStyle = cell.getCellStyle();
		var cellFont = arguments.workbook.getFontAt( cellStyle.getFontIndexAsInt() );
		var rgb = getColorHelper().getRGBFromCellFont( arguments.workbook, cellFont );
		return {
			alignment: cellStyle.getAlignment().toString()
			,bold: cellFont.getBold()
			,bottomborder: cellStyle.getBorderBottom().toString()
			,bottombordercolor: getColorHelper().getRgbTripletForStyleColorFormat( arguments.workbook, cellStyle, "bottombordercolor" )
			,color: ArrayToList( rgb )
			,dataformat: cellStyle.getDataFormatString()
			,fgcolor: getColorHelper().getRgbTripletForStyleColorFormat( arguments.workbook, cellStyle, "fgcolor" )
			,fillpattern: cellStyle.getFillPattern().toString()
			,font: cellFont.getFontName()
			,fontsize: cellFont.getFontHeightInPoints()
			,indent: cellStyle.getIndention()
			,italic: cellFont.getItalic()
			,leftborder: cellStyle.getBorderLeft().toString()
			,leftbordercolor: getColorHelper().getRgbTripletForStyleColorFormat( arguments.workbook, cellStyle, "leftbordercolor" )
			,quoteprefixed: cellStyle.getQuotePrefixed()
			,rightborder: cellStyle.getBorderRight().toString()
			,rightbordercolor: getColorHelper().getRgbTripletForStyleColorFormat( arguments.workbook, cellStyle, "rightbordercolor" )
			,rotation: cellStyle.getRotation()
			,strikeout: cellFont.getStrikeout()
			,textwrap: cellStyle.getWrapText()
			,topborder: cellStyle.getBorderTop().toString()
			,topbordercolor: getColorHelper().getRgbTripletForStyleColorFormat( arguments.workbook, cellStyle, "topbordercolor" )
			,underline: getFormatHelper().underlineNameFromIndex( cellFont.getUnderline() )
			,verticalalignment: cellStyle.getVerticalAlignment().toString()
		};
	}

	public any function getCellFormula( required workbook, numeric row, numeric column ){
		if( !arguments.KeyExists( "row" ) || !arguments.KeyExists( "column" ) )
			return getSheetHelper().getAllSheetFormulas( arguments.workbook );
		var cell = getCellHelper().getCellAt( arguments.workbook, arguments.row, arguments.column );
		if( IsNull( cell ) )
			return "";
		if( getCellHelper().cellIsOfType( cell, "FORMULA" ) )
			return cell.getCellFormula();
		return "";
	}

	public string function getCellHyperLink( required workbook, required numeric row, required numeric column ){
		var cell = getCellHelper().initializeCell( arguments.workbook, arguments.row, arguments.column );
		return cell.getHyperLink()?.getAddress()?:"";
	}

	public string function getCellType( required workbook, required numeric row, required numeric column ){
		var cell = getCellHelper().getCellAt( arguments.workbook, arguments.row, arguments.column );
		if( IsNull( cell ) )
			return "";
		return cell.getCellType().toString();
	}

	public any function getCellValue( required workbook, required numeric row, required numeric column, boolean returnVisibleValue=true ){
		var cell = getCellHelper().getCellAt( arguments.workbook, arguments.row, arguments.column );
		if( IsNull( cell ) )
			return "";
		if( arguments.returnVisibleValue && !getCellHelper().cellIsOfType( cell, "FORMULA" ) )
			return getCellHelper().getFormattedCellValue( cell );
		return getCellHelper().getCellValueAsType( arguments.workbook, cell );
	}

	public numeric function getColumnCount( required workbook, sheetNameOrNumber ){
		if( arguments.KeyExists( "sheetNameOrNumber" ) )
			getSheetHelper().setActiveSheetNameOrNumber( argumentCollection=arguments );
		var result = 0;
		var rowIterator = getSheetHelper().getActiveSheetRowIterator( arguments.workbook );
		while( rowIterator.hasNext() ){
			var row = rowIterator.next();
			result = Max( result, row.getLastCellNum() );
		}
		return result;
	}

	public numeric function getColumnWidth( required workbook, required numeric column ){
		var columnIndex = ( arguments.column -1 );
		return ( getSheetHelper().getActiveSheet( arguments.workbook ).getColumnWidth( JavaCast( "int", columnIndex ) ) / 256 );// whole character width (of zero character)
	}

	public numeric function getColumnWidthInPixels( required workbook, required numeric column ){
		var columnIndex = ( arguments.column -1 );
		return getSheetHelper().getActiveSheet( arguments.workbook ).getColumnWidthInPixels( JavaCast( "int", columnIndex ) );
	}

	public numeric function getLastRowNumber( required workbook, sheetNameOrNumber ){
		if( arguments.KeyExists( "sheetNameOrNumber" ) )
			getSheetHelper().setActiveSheetNameOrNumber( argumentCollection=arguments );
		var sheet = getSheetHelper().getActiveSheet( arguments.workbook );
		var lastRowIndex = getSheetHelper().getLastRowIndex( sheet );
		return lastRowIndex +1;
	}

	public array function getPresetColorNames(){
		var presetEnum = createJavaObject( "org.apache.poi.hssf.util.HSSFColor$HSSFColorPredefined" );
		var result = [];
		for( var value in presetEnum.values() )
			result.Append( value.name() );
		return result.Sort( "text" );
	}

	public boolean function getRecalculateFormulasOnNextOpen( required workbook, string sheetName ){
		if( arguments.KeyExists( "sheetName" ) ){
			var sheet = getSheetHelper().getSheetByName( arguments.workbook, arguments.sheetName );
			return sheet.getForceFormulaRecalculation();
		}
		if( isXmlFormat( arguments.workbook ) )
			return arguments.workbook.getForceFormulaRecalculation();
		// HSSF doesn't seem to work at the workbook level, so just decide based on the active sheet's flag
		return getSheetHelper().getActiveSheet( arguments.workbook ).getForceFormulaRecalculation();
	}

	public numeric function getRowCount( required workbook, sheetNameOrNumber ){
		return getLastRowNumber( argumentCollection=arguments );
	}

	public string function getSheetPrintOrientation( required workbook, string sheetName, numeric sheetNumber ){
		return getSheetHelper().getPrintOrientation( argumentCollection=arguments );
	}

	public SpreadSheet function groupColumns( required workbook, required numeric startColumn, required numeric endColumn ){
		getSheetHelper().getActiveSheet( arguments.workbook ).groupColumn( JavaCast( "int", arguments.startColumn-1 ), JavaCast( "int", arguments.endColumn-1 ) );
		return this;
	}

	public SpreadSheet function groupRows( required workbook, required numeric startRow, required numeric endRow ){
		getSheetHelper().getActiveSheet( arguments.workbook ).groupRow( JavaCast( "int", arguments.startRow-1 ), JavaCast( "int", arguments.endRow-1 ) );
		return this;
	}

	public Spreadsheet function hideColumn( required workbook, required numeric column ){
		getColumnHelper().toggleColumnHidden( arguments.workbook, arguments.column, true );
		return this;
	}

	public Spreadsheet function hideRow( required workbook, required numeric row ){
		getRowHelper().toggleRowHidden( arguments.workbook, arguments.row, true );
		return this;
	}

	public Spreadsheet function hideSheet( required workbook, string sheetName, numeric sheetNumber ){
		arguments.sheetNumber = getSheetHelper().getSheetNumberFromArguments( argumentCollection=arguments );
		getSheetHelper().setVisibility( arguments.workbook, arguments.sheetNumber, "HIDDEN" );
		return this;
	}

	public struct function info( required workbookOrPath ){
		/*
		properties returned in the struct are:
			* AUTHOR
			* CATEGORY
			* COMMENTS
			* CREATIONDATE
			* LASTEDITED
			* LASTAUTHOR
			* LASTSAVED
			* KEYWORDS
			* MANAGER
			* COMPANY
			* SUBJECT
			* TITLE
			* SHEETS
			* SHEETNAMES
			* SPREADSHEETTYPE
		*/
		 //use this.isSpreadsheetObject to avoid clash with ACF built-in function
		var workbook = this.isSpreadsheetObject( arguments[ 1 ] )? arguments[ 1 ]: getWorkbookHelper().workbookFromFile( arguments[ 1 ] );
		//format specific metadata
		var info = isBinaryFormat( workbook )? getInfoHelper().binaryInfo( workbook ): getInfoHelper().xmlInfo( workbook );
		//common properties
		info.sheets = workbook.getNumberOfSheets();
		var sheetnames = [];
		if( IsNumeric( info.sheets ) ){
			for( var i = 1; i <= info.sheets; i++ )
				sheetnames.Append( workbook.getSheetName( JavaCast( "int", ( i -1 ) ) ) );
			info.sheetnames = sheetnames.ToList();
		}
		info.spreadSheetType = isXmlFormat( workbook )? "Excel (2007)": "Excel";
		return info;
	}

	public boolean function isBinaryFormat( required workbook ){
		return arguments.workbook.getClass().getCanonicalName() == getClassHelper().getClassName( "HSSFWorkbook" );
	}

	public boolean function isColumnHidden( required workbook, required numeric column ){
		return getSheetHelper().getActiveSheet( arguments.workbook ).isColumnHidden( arguments.column - 1 );
	}

	public boolean function isRowHidden( required workbook, required numeric row ){
		var rowObject = getRowHelper().getRowFromActiveSheet( arguments.workbook, arguments.row );
		return getRowHelper().rowIsHidden( rowObject );
	}

	public boolean function isSpreadsheetFile( required string path ){
		getFileHelper().throwErrorIFfileNotExists( arguments.path );
		try{
			var workbook = getWorkbookHelper().workbookFromFile( arguments.path );
		}
		catch( cfsimplicity.spreadsheet.invalidFile exception ){
			return false;
		}
		return true;
	}

	public boolean function isSpreadsheetObject( required object ){
		return isBinaryFormat( arguments.object ) || isXmlFormat( arguments.object );
	}

	public boolean function isXmlFormat( required workbook ){
		//ACF doesn't support [].Find( needle ) in all contexts;
		return ArrayFind( [ getClassHelper().getClassName( "XSSFWorkbook" ), getClassHelper().getClassName( "SXSSFWorkbook" ) ], arguments.workbook.getClass().getCanonicalName() );
	}

	public boolean function isStreamingXmlFormat( required workbook ){
		return arguments.workbook.getClass().getCanonicalName() == getClassHelper().getClassName( "SXSSFWorkbook" );
	}

	public Spreadsheet function mergeCells(
		required workbook
		,required numeric startRow
		,required numeric endRow
		,required numeric startColumn
		,required numeric endColumn
		,boolean emptyInvisibleCells=false
	){
		if( arguments.startRow < 1 || arguments.startRow > arguments.endRow )
			Throw( type=this.getExceptionType() & ".invalidStartOrEndRowArgument", message="Invalid startRow or endRow", detail="Row values must be greater than 0 and the startRow cannot be greater than the endRow." );
		if( arguments.startColumn < 1 || arguments.startColumn > arguments.endColumn )
			Throw( type=this.getExceptionType() & ".invalidStartOrEndColumnArgument", message="Invalid startColumn or endColumn", detail="Column values must be greater than 0 and the startColumn cannot be greater than the endColumn." );
		var indices = {
			startRow: ( arguments.startRow - 1 )
			,endRow: ( arguments.endRow - 1 )
			,startColumn: ( arguments.startColumn - 1 )
			,endColumn: ( arguments.endColumn - 1 )
		};
		var cellRangeAddress = getRangeHelper().getCellRangeAddressFromColumnAndRowIndices( indices );
		getSheetHelper().getActiveSheet( arguments.workbook ).addMergedRegion( cellRangeAddress );
		if( !arguments.emptyInvisibleCells )
			return this;
		// stash the value to retain
		var visibleValue = getCellValue( arguments.workbook, arguments.startRow, arguments.startColumn );
		//empty all cells in the merged region
		setCellRangeValue( arguments.workbook, "", arguments.startRow, arguments.endRow, arguments.startColumn, arguments.endColumn );
		//restore the stashed value
		setCellValue( arguments.workbook, visibleValue, arguments.startRow, arguments.startColumn );
		return this;
	}

	public Spreadsheet function moveSheet( required workbook, required string sheetName, required numeric newPosition ){
		getSheetHelper()
			.validateSheetExistsWithName( arguments.workbook, arguments.sheetName )
			.validateSheetNumber( arguments.workbook, arguments.newPosition )
			.moveSheet( arguments.workbook, arguments.sheetName, arguments.newPosition-1 );
		return this;
	}

	public any function new(
		string sheetName="Sheet1"
		,boolean xmlFormat=( getDefaultWorkbookFormat() == "xml" )
		,boolean streamingXml=false
		,numeric streamingWindowSize
	){
		if( arguments.streamingXml && !arguments.xmlFormat )
			arguments.xmlFormat = true;
		var createArgs.type = getWorkbookHelper().typeFromArguments( arguments.xmlFormat, arguments.streamingXml );
		if( arguments.KeyExists( "streamingWindowSize" ) )
			createArgs.streamingWindowSize = arguments.streamingWindowSize;
		var workbook = getWorkbookHelper().createWorkBook( argumentCollection=createArgs );
		getSheetHelper().validateSheetName( arguments.sheetName );
		createSheet( workbook, arguments.sheetName, arguments.xmlFormat );
		return workbook;
	}

	public any function newChainable( existingWorkbookOrNewWorkbookType="" ){
		return New objects.SpreadsheetChainable( this, arguments.existingWorkbookOrNewWorkbookType );
	}

	public any function newStreamingXlsx( string sheetName="Sheet1", numeric streamingWindowSize=100 ){
		return new(
			sheetName=arguments.sheetName
			,xmlFormat=true
			,streamingXml=true
			,streamingWindowSize=arguments.streamingWindowSize
		);
	}

	public any function newConditionalFormatting(){
		return New objects.ConditionalFormatting( this );
	}

	public any function newDataValidation(){
		return New objects.DataValidation( this );
	}

	public any function newXls( string sheetName="Sheet1" ){
		return new( sheetName=arguments.sheetName, xmlFormat=false );
	}

	public any function newXlsx( string sheetName="Sheet1" ){
		return new( sheetName=arguments.sheetName, xmlFormat=true );
	}

	public any function processLargeFile( required string filepath ){
		return New objects.ProcessLargeFile( this, arguments.filepath );
	}

	/* row order is not guaranteed if using more than one thread */
	public string function queryToCsv( required query query, boolean includeHeaderRow=false, string delimiter=",", numeric threads=0 ){
		return writeCsv()
			.fromData( arguments.query )
			.withQueryColumnsAsHeader( arguments.includeHeaderRow )
			.withDelimiter( arguments.delimiter )
			.withParallelThreads( arguments.threads )
			.execute();
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
		if( arguments.KeyExists( "query" ) )
			Throw( type=this.getExceptionType() & ".invalidQueryArgument", message="Invalid argument 'query'.", detail="Just use format='query' to return a query object" );
		getExceptionHelper().throwExceptionIFreadFormatIsInvalid( argumentCollection=arguments );
		getSheetHelper().throwErrorIFSheetNameAndNumberArgumentsBothPassed( argumentCollection=arguments );
		getFileHelper().throwErrorIFfileNotExists( arguments.src );
		var passwordProtected = ( arguments.KeyExists( "password") && !arguments.password.Trim().IsEmpty() );
		var workbook = passwordProtected? getWorkbookHelper().workbookFromFile( arguments.src, arguments.password ): getWorkbookHelper().workbookFromFile( arguments.src );
		if( arguments.KeyExists( "sheetName" ) )
			setActiveSheet( workbook=workbook, sheetName=arguments.sheetName );
		if( !arguments.KeyExists( "format" ) )
			return workbook;
		var args = { workbook: workbook };
		if( arguments.KeyExists( "sheetName" ) )
			args.sheetName = arguments.sheetName;
		if( arguments.KeyExists( "sheetNumber" ) )
			args.sheetNumber = arguments.sheetNumber;
		if( arguments.KeyExists( "headerRow" ) ){
			args.headerRow = arguments.headerRow;
			args.includeHeaderRow = arguments.includeHeaderRow;
		}
		if( arguments.KeyExists( "rows" ) )
			args.rows = arguments.rows;
		if( arguments.KeyExists( "columns" ) )
			args.columns = arguments.columns;
		if( arguments.KeyExists( "columnNames" ) )
			args.columnNames = arguments.columnNames; // columnNames is what cfspreadsheet action="read" uses
		else if( arguments.KeyExists( "queryColumnNames" ) )
			args.columnNames = arguments.queryColumnNames;// accept better alias `queryColumnNames` to match csvToQuery
		if( ( arguments.format == "query" ) && arguments.KeyExists( "queryColumnTypes" ) ){
			args.queryColumnTypes = arguments.queryColumnTypes;
			getQueryHelper().throwErrorIFinvalidQueryColumnTypesArgument( argumentCollection=args );
		}
		args.includeBlankRows = arguments.includeBlankRows;
		args.fillMergedCellsWithVisibleValue = arguments.fillMergedCellsWithVisibleValue;
		args.includeHiddenColumns = arguments.includeHiddenColumns;
		args.includeHiddenRows = arguments.includeHiddenRows;
		args.includeRichTextFormatting = arguments.includeRichTextFormatting;
		args.makeColumnNamesSafe = arguments.makeColumnNamesSafe;
		args.returnVisibleValues = arguments.returnVisibleValues;
		if( arguments.format == "array" )
			return getSheetHelper().sheetToArray( argumentCollection=args );
		if( arguments.format == "arrayOfStructs" )
			return getSheetHelper().sheetToArrayOfStructs( argumentCollection=args );
		var generatedQuery = getSheetHelper().sheetToQuery( argumentCollection=args );
		if( arguments.format == "query" )
			return generatedQuery;
		if( arguments.format == "csv" ){
			return writeCsv()
				.fromData( generatedQuery )
				.withQueryColumnsAsHeader( arguments.includeHeaderRow )
				.withDelimiter( arguments.csvDelimiter )
				.execute();
		}
		// format = html
		return getQueryHelper().queryToHtml( generatedQuery, arguments.includeHeaderRow );
	}

	public any function readBinary( required workbook ){
		var baos = createJavaObject( "org.apache.commons.io.output.ByteArrayOutputStream" ).init();
		arguments.workbook.write( baos );
		baos.flush();
		return baos.toByteArray();
	}

	public any function readCsv( required string filepath ){
		return New objects.ReadCsv( this, arguments.filepath );
	}

	public any function readLargeFile(
		required string src
		,string format="query"
		,string sheetName
		,numeric sheetNumber // 1-based
		,numeric headerRow
		,boolean includeHeaderRow=false
		,boolean includeBlankRows=false
		,boolean includeHiddenColumns=true
		,boolean includeHiddenRows=true
		,any queryColumnNames //list or array
		,any queryColumnTypes //'auto', list of types, or struct of column names/types mapping. Null means no types are specified.
		,boolean makeColumnNamesSafe=false
		,string password
		,string csvDelimiter=","
		,struct streamingOptions
		,boolean returnVisibleValues=false
	){
		getFileHelper().throwErrorIFfileNotExists( arguments.src );
		getExceptionHelper().throwExceptionIFreadFormatIsInvalid( argumentCollection=arguments );
		getSheetHelper().throwErrorIFSheetNameAndNumberArgumentsBothPassed( argumentCollection=arguments );
		if( arguments.KeyExists( "streamingReaderOptions" ) ) //support legacy naming
			arguments.streamingOptions = arguments.streamingReaderOptions;
		var builderOptions = arguments.streamingReaderOptions?:{};
		if( arguments.KeyExists( "password" ) )
			builderOptions.password = arguments.password;
		var sheetToQueryArgs = {
			includeBlankRows: arguments.includeBlankRows
			,includeHiddenColumns: arguments.includeHiddenColumns
			,includeHiddenRows: arguments.includeHiddenRows
			,makeColumnNamesSafe: arguments.makeColumnNamesSafe
			,returnVisibleValues = arguments.returnVisibleValues
		};
		if( arguments.KeyExists( "sheetName" ) )
			sheetToQueryArgs.sheetName = arguments.sheetName;
		if( arguments.KeyExists( "sheetNumber" ) )
			sheetToQueryArgs.sheetNumber = arguments.sheetNumber;
		if( arguments.KeyExists( "headerRow" ) ){
			sheetToQueryArgs.headerRow = arguments.headerRow;
			sheetToQueryArgs.includeHeaderRow = arguments.includeHeaderRow;
		}
		if( arguments.KeyExists( "queryColumnNames" ) )
			sheetToQueryArgs.columnNames = arguments.queryColumnNames;
		if( ( arguments.format == "query" ) && arguments.KeyExists( "queryColumnTypes" ) ){
			sheetToQueryArgs.queryColumnTypes = arguments.queryColumnTypes;
			getQueryHelper().throwErrorIFinvalidQueryColumnTypesArgument( argumentCollection=sheetToQueryArgs );
		}
		var generatedQuery = getStreamingReaderHelper().readFileIntoQuery( arguments.src, builderOptions, sheetToQueryArgs );
		if( arguments.format == "query" )
			return generatedQuery;
		if( arguments.format == "csv" ){
			return writeCsv()
				.fromData( generatedQuery )
				.withQueryColumnsAsHeader( arguments.includeHeaderRow )
				.withDelimiter( arguments.csvDelimiter )
				.execute();
		}
		// format = html
		return getQueryHelper().queryToHtml( generatedQuery, arguments.includeHeaderRow );
	}

	public Spreadsheet function recalculateAllFormulas( required workbook ){
		arguments.workbook.getCreationHelper().createFormulaEvaluator().evaluateAllFormulaCells( arguments.workbook );
		return this;
	}

	public SpreadSheet function removeColumnBreak( required workbook, required numeric column ){
		getSheetHelper().getActiveSheet( arguments.workbook ).removeColumnBreak( JavaCast( "int", ( arguments.column-1 ) ) );
		return this;
	}

	public SpreadSheet function removeRowBreak( required workbook, required numeric row ){
		getSheetHelper().getActiveSheet( arguments.workbook ).removeRowBreak( JavaCast( "int", ( arguments.row-1 ) ) );
		return this;
	}

	public Spreadsheet function removePrintGridlines( required workbook ){
		getSheetHelper().getActiveSheet( arguments.workbook ).setPrintGridlines( JavaCast( "boolean", false ) );
		return this;
	}

	public Spreadsheet function removeSheet( required workbook, required string sheetName ){
		getSheetHelper()
			.validateSheetName( arguments.sheetName )
			.validateSheetExistsWithName( arguments.workbook, arguments.sheetName );
		arguments.sheetNumber = ( arguments.workbook.getSheetIndex( arguments.sheetName ) +1 );
		var sheetIndex = ( sheetNumber -1 );
		getSheetHelper().deleteSheetAtIndex( arguments.workbook, sheetIndex );
		return this;
	}

	public Spreadsheet function removeSheetNumber( required workbook, required numeric sheetNumber ){
		getSheetHelper().validateSheetNumber( arguments.workbook, arguments.sheetNumber );
		var sheetIndex = ( arguments.sheetNumber -1 );
		getSheetHelper().deleteSheetAtIndex( arguments.workbook, sheetIndex );
		return this;
	}

	public Spreadsheet function renameSheet( required workbook, required string sheetName, required numeric sheetNumber ){
		getSheetHelper()
			.validateSheetName( arguments.sheetName )
			.validateSheetNumber( arguments.workbook, arguments.sheetNumber );
		var sheetIndex = ( arguments.sheetNumber -1 );
		var foundAt = arguments.workbook.getSheetIndex( JavaCast( "string", arguments.sheetName ) );
		if( ( foundAt > 0 ) && ( foundAt != sheetIndex ) )
			Throw( type=this.getExceptionType() & ".invalidSheetName", message="Invalid Sheet Name [#arguments.sheetName#]", detail="The workbook already contains a sheet named [#sheetName#]. Sheet names must be unique" );
		arguments.workbook.setSheetName( JavaCast( "int", sheetIndex ), JavaCast( "string", arguments.sheetName ) );
		return this;
	}

	public Spreadsheet function setActiveCell( required workbook, required numeric row, required numeric column ){
		var sheet = getSheetHelper().getActiveSheet( arguments.workbook );
		var cell = getCellHelper().initializeCell( arguments.workbook, arguments.row, arguments.column );
		var cellAddress = createJavaObject( "org.apache.poi.ss.util.CellAddress" ).init( cell );
		sheet.setActiveCell( cellAddress );
		return this;
	}

	public Spreadsheet function setActiveSheet( required workbook, string sheetName, numeric sheetNumber ){
		getSheetHelper().validateSheetNameOrNumberWasProvided( argumentCollection=arguments );
		if( arguments.KeyExists( "sheetName" ) ){
			getSheetHelper().validateSheetExistsWithName( arguments.workbook, arguments.sheetName );
			arguments.sheetNumber = ( arguments.workbook.getSheetIndex( JavaCast( "string", arguments.sheetName ) ) + 1 );
		}
		getSheetHelper().validateSheetNumber( arguments.workbook, arguments.sheetNumber );
		arguments.workbook.setActiveSheet( JavaCast( "int", ( arguments.sheetNumber - 1 ) ) );
		return this;
	}

	public Spreadsheet function setActiveSheetNumber( required workbook, numeric sheetNumber ){
		setActiveSheet( workbook=arguments.workbook, sheetNumber=arguments.sheetNumber );
		return this;
	}

	public Spreadsheet function setCellComment(
		required workbook
		,required struct comment
		,required numeric row
		,required numeric column
	){
		/*
		The comment struct may contain the following keys:
			* anchor
			* author
			* bold
			* color
			* comment
			* fillcolor
			* font
			* horizontalalignment
			* italic
			* linestyle
			* linestylecolor
			* size
			* strikeout
			* underline
			* verticalalignment
			* visible
		 */
		var factory = arguments.workbook.getCreationHelper();
		var commentString = factory.createRichTextString( JavaCast( "string", arguments.comment.comment ) );
		var cellAddress = { row: arguments.row, column: arguments.column };
		var anchor = getCommentHelper().createAnchor( factory, arguments.comment, cellAddress );
		var drawingPatriarch = getSheetHelper().getActiveSheet( arguments.workbook ).createDrawingPatriarch();
		var commentObject = drawingPatriarch.createCellComment( anchor );
		if( arguments.comment.KeyExists( "author" ) )
			commentObject.setAuthor( JavaCast( "string", arguments.comment.author ) );
		if( arguments.comment.KeyExists( "visible" ) )
			commentObject.setVisible( JavaCast( "boolean", arguments.comment.visible ) );//doesn't always seem to work
		getCommentHelper().addFontStylesToComment( arguments.comment, arguments.workbook, commentString );
		if( isBinaryFormat( arguments.workbook ) )
			getCommentHelper().addHSSFonlyStyles( arguments.comment, commentObject );
		commentObject.setString( commentString );
		var cell = getCellHelper().initializeCell( arguments.workbook, arguments.row, arguments.column );
		cell.setCellComment( commentObject );
		return this;
	}

	public Spreadsheet function setCellFormula(
		required workbook
		,required string formula
		,required numeric row
		,required numeric column
	){
		var cell = getCellHelper().initializeCell( arguments.workbook, arguments.row, arguments.column );
		cell.setCellFormula( JavaCast( "string", arguments.formula ) );
		if( getReturnCachedFormulaValues() )
			getCellHelper().forceFormulaEvaluation( arguments.workbook, cell );
		return this;
	}

	public Spreadsheet function setCellHyperlink(
		required workbook
		,required string link
		,required numeric row
		,required numeric column
		,any cellValue
		,string type="URL"
		,any format={ color: "0000ff", underline: true } //struct or cellStyle object
		,string tooltip //xlsx only, maybe MS Excel full version only
	){
		arguments.type = arguments.type.UCase();
		getHyperLinkHelper().throwErrorIfTypeIsInvalid( arguments.type );
		getHyperLinkHelper().throwErrorIfTooltipAndWorkbookIsXls( argumentCollection=arguments );
		var cell = getCellHelper().initializeCell( arguments.workbook, arguments.row, arguments.column );
		getHyperLinkHelper().addHyperLinkToCell( cell=cell, argumentCollection=arguments );
		if( arguments.KeyExists( "cellValue" ) )
			getCellHelper().setCellValueAsType( arguments.workbook, cell, arguments.cellValue );
		formatCell( arguments.workbook, arguments.format, arguments.row, arguments.column );
		return this;
	}

	public Spreadsheet function setCellRangeValue(
		required workbook
		,required value
		,required numeric startRow
		,required numeric endRow
		,required numeric startColumn
		,required numeric endColumn
	){
		for( var rowNumber = arguments.endRow; rowNumber >= arguments.startRow; rowNumber-- ){
			for( var columnNumber = arguments.endColumn; columnNumber >= arguments.startColumn; columnNumber-- )
				setCellValue( arguments.workbook, arguments.value, rowNumber, columnNumber );
		}
		return this;
	}

	public Spreadsheet function setCellValue( required workbook, required value, required numeric row, required numeric column, string datatype ){
		var args = {
			workbook: arguments.workbook
			,cell: getCellHelper().initializeCell( arguments.workbook, arguments.row, arguments.column )
			,value: arguments.value
		};
		if( arguments.KeyExists( "datatype" ) )
			args.type = arguments.datatype;
		//support legacy argument name
		if( arguments.KeyExists( "type" ) )
			args.type = arguments.type;
		getCellHelper().setCellValueAsType( argumentCollection=args );
		return this;
	}

	public Spreadsheet function setColumnWidth( required workbook, required numeric column, required numeric width ){
		var columnIndex = ( arguments.column -1 );
		getSheetHelper().getActiveSheet( arguments.workbook ).setColumnWidth( JavaCast( "int", columnIndex ), JavaCast( "int", ( arguments.width * 256 ) ) );
		return this;
	}

	public Spreadsheet function setDateFormats( required struct dateFormats ){
		getDateHelper().setCustomFormats( arguments.dateFormats );
		return this;
	}

	public Spreadsheet function setFitToPage( required workbook, required boolean state, numeric pagesWide, numeric pagesHigh ){
		var sheet = getSheetHelper().getActiveSheet( arguments.workbook );
		sheet.setFitToPage( JavaCast( "boolean", arguments.state ) );
		sheet.setAutoBreaks( JavaCast( "boolean", arguments.state ) ); //seems dependent on this matching
		if( !arguments.state )
			return this;
		if( arguments.KeyExists( "pagesWide" ) && IsValid( "integer", arguments.pagesWide ) )
			sheet.getPrintSetup().setFitWidth( JavaCast( "short", arguments.pagesWide ) );
		if( arguments.KeyExists( "pagesWide" ) && IsValid( "integer", arguments.pagesHigh ) )
			sheet.getPrintSetup().setFitHeight( JavaCast( "short", arguments.pagesHigh ) );
		return this;
	}

	public Spreadsheet function setFooter(
		required workbook
		,string leftFooter=""
		,string centerFooter=""
		,string rightFooter=""
	){
		var footer = getSheetHelper().getActiveSheetFooter( arguments.workbook );
		if( arguments.centerFooter.Len() )
			footer.setCenter( JavaCast( "string", arguments.centerFooter ) );
		if( arguments.leftFooter.Len() )
			footer.setleft( JavaCast( "string", arguments.leftFooter ) );
		if( arguments.rightFooter.Len() )
			footer.setright( JavaCast( "string", arguments.rightFooter ) );
		return this;
	}

	public Spreadsheet function setFooterImage(
		required workbook
		,required string position // left|center|right
		,required any image
		,string imageType
	){
		getHeaderImageHelper().setHeaderOrFooterImage( argumentCollection=arguments, isHeader=false );
		return this;
	}

	public Spreadsheet function setHeader(
		required workbook
		,string leftHeader=""
		,string centerHeader=""
		,string rightHeader=""
	){
		var header = getSheetHelper().getActiveSheetHeader( arguments.workbook );
		if( arguments.centerHeader.Len() )
			header.setCenter( JavaCast( "string", arguments.centerHeader ) );
		if( arguments.leftHeader.Len() )
			header.setleft( JavaCast( "string", arguments.leftHeader ) );
		if( arguments.rightHeader.Len() )
			header.setright( JavaCast( "string", arguments.rightHeader ) );
		return this;
	}

	public Spreadsheet function setHeaderImage(
		required workbook
		,required string position // left|center|right
		,required any image
		,string imageType
	){
		getHeaderImageHelper().setHeaderOrFooterImage( argumentCollection=arguments );
		return this;
	}

	public Spreadsheet function setReadOnly( required workbook, required string password ){
		if( isXmlFormat( arguments.workbook ) )
			Throw( type=this.getExceptionType() & ".invalidSpreadsheetType", message="setReadOnly not supported for XML workbooks", detail="The setReadOnly() method only works on binary 'xls' workbooks." );
		// writeProtectWorkbook takes both a user name and a password, just making up a user name
		arguments.workbook.writeProtectWorkbook( JavaCast( "string", arguments.password ), JavaCast( "string", "user" ) );
		return this;
	}

	public Spreadsheet function setRecalculateFormulasOnNextOpen( required workbook, boolean value=true, string sheetName ){
		if( arguments.KeyExists( "sheetName" ) ){
			var sheet = getSheetHelper().getSheetByName( arguments.workbook, arguments.sheetName );
			sheet.setForceFormulaRecalculation( JavaCast( "boolean", arguments.value ) );
			return this;
		}
		arguments.workbook.setForceFormulaRecalculation( JavaCast( "boolean", arguments.value ) );
		if( isXmlFormat( arguments.workbook ) )
			return this;
		// HSSF setForceFormulaRecalculation() doesn't seem to work at the workbook level, so set all sheets to recalculate
		var sheetIterator = arguments.workbook.sheetIterator();
		while( sheetIterator.hasNext() ){
			var sheet = sheetIterator.next();
			sheet.setForceFormulaRecalculation( JavaCast( "boolean", arguments.value ) );
		}
		return this;
	}

	public Spreadsheet function setRepeatingColumns( required workbook, required string columnRange ){
		arguments.columnRange = arguments.columnRange.Trim();
		if( !IsValid( "regex", arguments.columnRange, "[A-Za-z]:[A-Za-z]" ) )
			Throw( type=this.getExceptionType() & ".invalidColumnRangeArgument", message="Invalid columnRange argument", detail="The 'columnRange' argument should be in the form 'A:B'" );
		var cellRangeAddress = getRangeHelper().getCellRangeAddressFromReference( arguments.columnRange );
		getSheetHelper().getActiveSheet( arguments.workbook ).setRepeatingColumns( cellRangeAddress );
		return this;
	}

	public Spreadsheet function setRepeatingRows( required workbook, required string rowRange ){
		arguments.rowRange = arguments.rowRange.Trim();
		if( !IsValid( "regex", arguments.rowRange,"\d+:\d+" ) )
			Throw( type=this.getExceptionType() & ".invalidRowRangeArgument", message="Invalid rowRange argument", detail="The 'rowRange' argument should be in the form 'n:n', e.g. '1:5'" );
		var cellRangeAddress = getRangeHelper().getCellRangeAddressFromReference( arguments.rowRange );
		getSheetHelper().getActiveSheet( arguments.workbook ).setRepeatingRows( cellRangeAddress );
		return this;
	}

	public SpreadSheet function setColumnBreak( required workbook, required numeric column ){
		getSheetHelper().getActiveSheet( arguments.workbook ).setColumnBreak( JavaCast( "int", ( arguments.column-1 ) ) );
		return this;
	}

	public SpreadSheet function setRowBreak( required workbook, required numeric row ){
		getSheetHelper().getActiveSheet( arguments.workbook ).setRowBreak( JavaCast( "int", ( arguments.row-1 ) ) );
		return this;
	}

	public Spreadsheet function setRowHeight( required workbook, required numeric row, required numeric height ){
		var rowObject = getRowHelper().getRowFromActiveSheet( arguments.workbook, arguments.row );
		if( IsNull( rowObject ) )
			getExceptionHelper().throwNonExistentRowException( arguments.row );
		rowObject.setHeightInPoints( JavaCast( "int", arguments.height ) );
		return this;
	}

	public Spreadsheet function setSheetTopMargin( required workbook, required numeric marginSize, string sheetName, numeric sheetNumber ){
		var sheet = getSheetHelper().getSpecifiedOrActiveSheet( argumentCollection=arguments );
		sheet.setMargin( sheet.TopMargin, arguments.marginSize );
		return this;
	}

	public Spreadsheet function setSheetBottomMargin( required workbook, required numeric marginSize, string sheetName, numeric sheetNumber ){
		var sheet = getSheetHelper().getSpecifiedOrActiveSheet( argumentCollection=arguments );
		sheet.setMargin( sheet.BottomMargin, arguments.marginSize );
		return this;
	}

	public Spreadsheet function setSheetLeftMargin( required workbook, required numeric marginSize, string sheetName, numeric sheetNumber ){
		var sheet = getSheetHelper().getSpecifiedOrActiveSheet( argumentCollection=arguments );
		sheet.setMargin( sheet.LeftMargin, arguments.marginSize );
		return this;
	}

	public Spreadsheet function setSheetRightMargin( required workbook, required numeric marginSize, string sheetName, numeric sheetNumber ){
		var sheet = getSheetHelper().getSpecifiedOrActiveSheet( argumentCollection=arguments );
		sheet.setMargin( sheet.RightMargin, arguments.marginSize );
		return this;
	}

	public Spreadsheet function setSheetHeaderMargin( required workbook, required numeric marginSize, string sheetName, numeric sheetNumber ){
		var sheet = getSheetHelper().getSpecifiedOrActiveSheet( argumentCollection=arguments );
		sheet.setMargin( sheet.HeaderMargin, arguments.marginSize );
		return this;
	}

	public Spreadsheet function setSheetFooterMargin( required workbook, required numeric marginSize, string sheetName, numeric sheetNumber ){
		var sheet = getSheetHelper().getSpecifiedOrActiveSheet( argumentCollection=arguments );
		sheet.setMargin( sheet.FooterMargin, arguments.marginSize );
		return this;
	}

	public Spreadsheet function setSheetPrintOrientation( required workbook, required string mode, string sheetName, numeric sheetNumber ){
		if( !ListFindNoCase( "landscape,portrait", arguments.mode ) )
			Throw( type=this.getExceptionType() & ".invalidModeArgument", message="Invalid mode argument", detail="#mode# is not a valid 'mode' argument. Use 'portrait' or 'landscape'" );
		var setToLandscape = ( LCase( arguments.mode ) == "landscape" );
		var sheet = getSheetHelper().getSpecifiedOrActiveSheet( argumentCollection=arguments );
		sheet.getPrintSetup().setLandscape( JavaCast( "boolean", setToLandscape ) );
		return this;
	}

	public struct function sheetInfo( required workbook, numeric sheetNumber ){
		return getSheetHelper().info( argumentCollection=arguments );
	}

	public Spreadsheet function shiftColumns( required workbook, required numeric start, numeric end=arguments.start, numeric offset=1 ){
		/*
			20210427 POI 4.x's sheet.shiftColumns() doesn't seem to work reliably: XSSF version doesn't delete columns that should be replaced. Both result in errors when writing
		*/
		if( arguments.start <= 0 )
			Throw( type=this.getExceptionType() & ".invalidStartArgument", message="Invalid start value", detail="The start value must be greater than or equal to 1" );
		if( arguments.KeyExists( "end" ) && ( ( arguments.end <= 0 ) || ( arguments.end < arguments.start ) ) )
			Throw( type=this.getExceptionType() & ".invalidEndArgument", message="Invalid end value", detail="The end value must be greater than or equal to the start value" );
		var rowIterator = getSheetHelper().getActiveSheetRowIterator( arguments.workbook );
		var startIndex = ( arguments.start -1 );
		var endIndex = arguments.KeyExists( "end" )? ( arguments.end -1 ): startIndex;
		while( rowIterator.hasNext() ){
			var row = rowIterator.next();
			if( arguments.offset > 0 ){
				for( var i = endIndex; i >= startIndex; i-- )
					getCellHelper().shiftCell( arguments.workbook, row, i, arguments.offset );
			}
			else{
				for( var i = startIndex; i <= endIndex; i++ )
					getCellHelper().shiftCell( arguments.workbook, row, i, arguments.offset );
			}
		}
		return this;
	}

	public Spreadsheet function shiftRows( required workbook, required numeric start, numeric end=arguments.start, numeric offset=1 ){
		getSheetHelper().getActiveSheet( arguments.workbook ).shiftRows(
			JavaCast( "int", ( arguments.start - 1 ) )
			,JavaCast( "int", ( arguments.end - 1 ) )
			,JavaCast( "int", arguments.offset )
		);
		return this;
	}

	public Spreadsheet function showColumn( required workbook, required numeric column ){
		getColumnHelper().toggleColumnHidden( arguments.workbook, arguments.column, false );
		return this;
	}

	public Spreadsheet function showRow( required workbook, required numeric row ){
		getRowHelper().toggleRowHidden( arguments.workbook, arguments.row, false );
		return this;
	}

	public SpreadSheet function ungroupColumns( required workbook, required numeric startColumn, required numeric endColumn ){
		getSheetHelper().getActiveSheet( arguments.workbook ).ungroupColumn( JavaCast( "int", arguments.startColumn-1 ), JavaCast( "int", arguments.endColumn-1 ) );
		return this;
	}

	public SpreadSheet function ungroupRows( required workbook, required numeric startRow, required numeric endRow ){
		getSheetHelper().getActiveSheet( arguments.workbook ).ungroupRow( JavaCast( "int", arguments.startRow-1 ), JavaCast( "int", arguments.endRow-1 ) );
		return this;
	}

	public Spreadsheet function unhideSheet( required workbook, string sheetName, numeric sheetNumber ){
		arguments.sheetNumber = getSheetHelper().getSheetNumberFromArguments( argumentCollection=arguments );
		getSheetHelper().setVisibility( arguments.workbook, arguments.sheetNumber, "VISIBLE" );
		return this;
	}

	public any function workbookFromCsv(
		string csv
		,string filepath
		,boolean firstRowIsHeader=false
		,boolean boldHeaderRow=true
		,boolean trim=true
		,boolean xmlFormat=false
		,string delimiter
	){
		var conversionArgs = {
			firstRowIsHeader: arguments.firstRowIsHeader
			,trim: arguments.trim
		};
		if( arguments.KeyExists( "csv" ) )
			conversionArgs.csv = arguments.csv;
		if( arguments.KeyExists( "filepath" ) )
			conversionArgs.filepath = arguments.filepath;
		if( arguments.KeyExists( "delimiter" ) )
			conversionArgs.delimiter = arguments.delimiter;
		var data = csvToQuery( argumentCollection=conversionArgs );
		return workbookFromQuery(
			data=data
			,addHeaderRow=arguments.firstRowIsHeader
			,boldHeaderRow=arguments.boldHeaderRow
			,xmlFormat=arguments.xmlFormat
		);
	}

	public any function workbookFromQuery(
		required query data
		,boolean addHeaderRow=true
		,boolean boldHeaderRow=true
		,boolean xmlFormat=false
		,boolean streamingXml=false
		,numeric streamingWindowSize=100
		,boolean ignoreQueryColumnDataTypes=false
		,struct datatypes
		,boolean autoSizeColumns=false
	){
		var workbook = new( xmlFormat=arguments.xmlFormat, streamingXml=arguments.streamingXml, streamingWindowSize=arguments.streamingWindowSize );
		var addRowsArgs = {
			workbook: workbook
			,data: arguments.data
			,ignoreQueryColumnDataTypes: arguments.ignoreQueryColumnDataTypes
			,autoSizeColumns: arguments.autoSizeColumns
		};
		if( arguments.KeyExists( "datatypes" ) )
			addRowsArgs.datatypes = arguments.datatypes;
		if( arguments.addHeaderRow ){
			var columns = getQueryHelper()._QueryColumnArray( arguments.data );
			addRow( workbook, columns );
			if( arguments.boldHeaderRow )
				formatRow( workbook, { bold: true }, 1 );
			addRowsArgs.row = 2;
			addRowsArgs.column = 1;
		}
		addRows( argumentCollection=addRowsArgs );
		return workbook;
	}

	public Spreadsheet function write(
		required workbook
		,required string filepath
		,boolean overwrite=false
		,string password
		,string algorithm="agile"
	){
		if( !arguments.overwrite && FileExists( arguments.filepath ) )
			getExceptionHelper().throwFileExistsException( arguments.filepath );
		var passwordProtect = ( arguments.KeyExists( "password" ) && !arguments.password.Trim().IsEmpty() );
		if( passwordProtect && isBinaryFormat( arguments.workbook ) )
			Throw( type=this.getExceptionType() & ".invalidSpreadsheetType", message="Whole file password protection is not supported for binary workbooks", detail="Password protection only works with XML ('xlsx') workbooks." );
		try{
			lock name="#arguments.filepath#" timeout=5{
				var outputStream = CreateObject( "java", "java.io.FileOutputStream" ).init( arguments.filepath );
				arguments.workbook.write( outputStream );
				outputStream.flush();
			}
		}
		finally{
			// always close the stream. otherwise file may be left in a locked state if an unexpected error occurs
			getFileHelper().closeLocalFileOrStream( local, "outputStream" );
			cleanUpStreamingXml( arguments.workbook );
		}
		if( passwordProtect )
			getFileHelper().encryptFile( arguments.filepath, arguments.password, arguments.algorithm );
		return this;
	}

	public any function writeCsv(){
		return New objects.WriteCsv( this );
	}

	public Spreadsheet function writeFileFromQuery(
		required query data
		,required string filepath
		,boolean overwrite=false
		,boolean addHeaderRow=true
		,boolean boldHeaderRow=true
		,boolean xmlFormat=false
		,boolean streamingXml=false
		,numeric streamingWindowSize=100
		,boolean ignoreQueryColumnDataTypes=false
		,struct datatypes
	){
		if( !arguments.xmlFormat && ( ListLast( arguments.filepath, "." ) == "xlsx" ) )
			arguments.xmlFormat = true;
		var workbookFromQueryArgs = {
			data: arguments.data
			,addHeaderRow: arguments.addHeaderRow
			,boldHeaderRow: arguments.boldHeaderRow
			,xmlFormat: arguments.xmlFormat
			,streamingXml: arguments.streamingXml
			,streamingWindowSize: arguments.streamingWindowSize
			,ignoreQueryColumnDataTypes: arguments.ignoreQueryColumnDataTypes
		};
		if( arguments.KeyExists( "datatypes" ) )
			workbookFromQueryArgs.datatypes = arguments.datatypes;
		var workbook = workbookFromQuery( argumentCollection=workbookFromQueryArgs );
		// force to .xlsx if appropriate
		if( arguments.xmlFormat && ( ListLast( arguments.filepath, "." ) == "xls" ) )
			arguments.filepath &= "x";
		write( workbook=workbook, filepath=arguments.filepath, overwrite=arguments.overwrite );
		return this;
	}

	public Spreadsheet function writeToCsv(
		required workbook
		,required string filepath
		,boolean overwrite=false
		,string delimiter=","
		,boolean includeHeaderRow=true
		,numeric headerRow=1
	){
		if( !arguments.overwrite && FileExists( arguments.filepath ) )
			getExceptionHelper().throwFileExistsException( arguments.filepath );
		var data = getSheetHelper().sheetToQuery(
			workbook=arguments.workbook
			,headerRow=arguments.headerRow
			,includeHeaderRow=arguments.includeHeaderRow
			,makeColumnNamesSafe=true //doesn't affect the output: avoids ACF clunky workaround in _QueryNew()
		);
		writeCsv()
			.fromData( data )
			.toFile( arguments.filepath )
			.withDelimiter( arguments.delimiter )
			.execute();
		return this;
	}

}