component{

	property name="cellRange" default="";
	property name="validValues" type="array";
	property name="valuesSourceSheetName" default="";
	property name="valuesSourceCellRange" default="";
	property name="errorMessage" default="";
	property name="errorTitle" default="";
	property name="suppressDropdown" type="boolean" default="false";
	/* Internal */
	property name="library";
	property name="workbookIsXlsx" type="boolean";
	/* POI java objects */
	property name="workbook";
	property name="sheet";
	property name="cellRangeAddress";
	property name="dataValidationHelper";
	property name="dataValidation";
	property name="validationConstraint";

	public DataValidation function init( required spreadsheetLibrary ){
		variables.library = arguments.spreadsheetLibrary;
		return this;
	}

	/* Public builder API */

	public DataValidation function onCells( required string cellRange ){
		variables.cellRange = Trim( arguments.cellRange );
		return this;
	}

	public DataValidation function withValues( required array values ){
		variables.validValues = arguments.values;
		return this;
	}

	public DataValidation function withValuesFromSheetName( required string sheetName ){
		variables.valuesSourceSheetName = arguments.sheetName;
		return this;
	}

	public DataValidation function withValuesFromCells( required string cellRange ){
		variables.valuesSourceCellRange = arguments.cellRange;
		return this;
	}

	public DataValidation function withErrorTitle( required string errorTitle ){
		variables.errorTitle = arguments.errorTitle;
		return this;
	}

	public DataValidation function withErrorMessage( required string errorMessage ){
		variables.errorMessage = arguments.errorMessage;
		return this;
	}

	public DataValidation function withNoDropdownArrow(){
		variables.suppressDropdown = true;
		return this;
	}

	public DataValidation function addToWorkbook( required workbook ){
		variables.workbookIsXlsx = variables.library.isXmlFormat( arguments.workbook );
		variables.workbook = arguments.workbook;
		variables.sheet = variables.library.getSheetHelper().getActiveSheet( variables.workbook );
		variables.dataValidationHelper = variables.sheet.getDataValidationHelper();
		variables.cellRangeAddress = variables.library.getRangeHelper().getCellRangeAddressFromReference( variables.cellRange );
		var addressList = variables.library.getClassHelper().loadClass( "org.apache.poi.ss.util.CellRangeAddressList" );
		addressList.addCellRangeAddress( cellRangeAddress );
		// passed array will trump values in other cells if both provided
		variables.validationConstraint = IsArray( variables.validValues?:"" )? getConstraintFromArray(): getConstraintFromCells();
		variables.dataValidation = variables.dataValidationHelper.createValidation( validationConstraint, addressList );
		if( variables.workbookIsXlsx )
			variables.dataValidation.setShowErrorBox( JavaCast( "boolean", true ) );//required to enforce validation in XSSF
		if( variables.errorTitle.Len() || variables.errorMessage.Len() )
			variables.dataValidation.createErrorBox( variables.errorTitle, variables.errorMessage );
		if( variables.suppressDropdown )
			setDropdownSuppression();
		variables.sheet.addValidationData( variables.dataValidation );
		return this;
	}

	/* For testing */
	public string function targetCellRangeAppliedToSheet(){
		if( !sheetHasValidation() )
			return "";
		return validationAppliedToSheet().getRegions().getCellRangeAddress( 0 ).formatAsString();
	}

	public array function validValueArrayAppliedToSheet(){
		if( !sheetHasValidation() )
			return [];
		return constraintAppliedToSheet().getExplicitListValues();
	}

	public string function sourceCellsReferenceAppliedToSheet(){
		if( !sheetHasValidation() )
			return "";
		return constraintAppliedToSheet().getFormula1();
	}

	public string function errorTitleAppliedToSheet(){
		if( !sheetHasValidation() )
			return "";
		return validationAppliedToSheet().getErrorBoxTitle();
	}

	public string function errorMessageAppliedToSheet(){
		if( !sheetHasValidation() )
			return "";
		return validationAppliedToSheet().getErrorBoxText();
	}

	public boolean function suppressDropdownSettingArrowAppliedToSheet(){
		return validationAppliedToSheet().getSuppressDropDownArrow();
	}

	/* Private  */

	private any function getConstraintFromArray(){
		return variables.dataValidationHelper.createExplicitListConstraint( variables.validValues );
	}

	private any function getConstraintFromCells(){
		var sheetName = determineSourceSheetName();
		var cellReference = variables.library.getRangeHelper().convertRangeReferenceToAbsoluteAddress( variables.valuesSourceCellRange );
		var sheetAndCellReference = sheetName & "!" & cellReference;
		return variables.dataValidationHelper.createFormulaListConstraint( sheetAndCellReference );
	}

	private string function determineSourceSheetName(){
		if( variables.valuesSourceSheetName.Len() ){
			variables.library.getSheetHelper().validateSheetExistsWithName( variables.workbook, variables.valuesSourceSheetName );
			return variables.valuesSourceSheetName;
		}
		return variables.sheet.getSheetName();
	}

	private void function setDropdownSuppression(){
		// XSSFDataValidation requires explicitly setting suppression to FALSE in order to suppress the dropdown (WTF)!! 
		// See https://poi.apache.org/components/spreadsheet/quick-guide.html#Validation
		var falseForXlsxTrueForXls = !variables.workbookIsXlsx;
		variables.dataValidation.setSuppressDropDownArrow( JavaCast( "boolean", falseForXlsxTrueForXls ) );
	}

	private boolean function sheetHasValidation(){
		return ( variables.sheet.getDataValidations().Len() > 0 );
	}

	private any function validationAppliedToSheet(){
		return variables.sheet.getDataValidations()[ 1 ];
	}

	private any function constraintAppliedToSheet(){
		return validationAppliedToSheet().getValidationConstraint();
	}

}