component{

	property name="cellRange" default="";
	property name="validValues" type="array";
	property name="valuesSourceSheetName" default="";
	property name="valuesSourceCellRange" default="";
	property name="minDate" default="";
	property name="maxDate" default="";
	property name="minInteger" default="";
	property name="maxInteger" default="";
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

	public DataValidation function withMinDate( required string date ){
		variables.minDate = arguments.date;
		return this;
	}

	public DataValidation function withMaxDate( required string date ){
		variables.maxDate = arguments.date;
		return this;
	}

	public DataValidation function withMinInteger( required string value ){
		variables.minInteger = arguments.value;
		return this;
	}

	public DataValidation function withMaxInteger( required string value ){
		variables.maxInteger = arguments.value;
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
		var addressList = variables.library.createJavaObject( "org.apache.poi.ss.util.CellRangeAddressList" );
		addressList.addCellRangeAddress( cellRangeAddress );
		createConstraint();
		variables.dataValidation = variables.dataValidationHelper.createValidation( variables.validationConstraint, addressList );
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

	public string function getConstraintType(){
		switch( variables.validationConstraint.ValidationType?:0 ){
			case 1: return "integer";
			case 2: return "decimal";
			case 3: return "list";
			case 4: return "date";
			case 5: return "time";
			case 6: return "length";
			case 7: return "formula";
		}
		return "undefined";
	}

	public string function getConstraintOperator(){
		//Don't use Elvis operator to set default: ACF always treats the getOperator() integer as null for some reason
		switch( variables.validationConstraint.getOperator() ){
			case 0: return "BETWEEN";
			case 1: return "NOT_BETWEEN";
			case 2: return "EQUAL";
			case 3: return "NOT_EQUAL";
			case 4: return "GREATER_THAN";
			case 5: return "LESS_THAN";
			case 6: return "GREATER_OR_EQUAL";
			case 7: return "LESS_OR_EQUAL";
		}
		return "undefined";
	}

	/* Private  */
	private void function createConstraint(){
		// set the type from the configured variables
		// passed array will trump values in other cells if both provided
		if( IsArray( variables.validValues?:"" ) )
			return createListConstraintFromArray();
		if( variables.valuesSourceCellRange.Len() )
			return createListConstraintFromCells();
		if( Len( variables.minDate ) || Len( variables.maxDate ) )
			return createDateConstraint();
		if( Len( variables.minInteger ) || Len( variables.maxInteger ) )
			return createIntegerConstraint();
	}

	private void function createDateConstraint(){
		if( !Len( variables.maxDate ) ){
			variables.validationConstraint = variables.dataValidationHelper.createDateConstraint(
				getConstraintOperatorObject( "GREATER_OR_EQUAL" )
				,getWorkbookSpecificDateValue( variables.minDate )
				,JavaCast( "null", "" )
				,"yyyy-MM-dd" //xlsx ignores this? https://stackoverflow.com/a/44312964/204620
			);
			return;
		}
		if( !Len( variables.minDate ) ){
			variables.validationConstraint = variables.dataValidationHelper.createDateConstraint(
				getConstraintOperatorObject( "LESS_OR_EQUAL" )
				,getWorkbookSpecificDateValue( variables.maxDate )
				,JavaCast( "null", "" )
				,"yyyy-MM-dd"
			);
			return;
		}
		variables.validationConstraint = variables.dataValidationHelper.createDateConstraint(
			getConstraintOperatorObject( "BETWEEN" )
			,getWorkbookSpecificDateValue( variables.minDate )
			,getWorkbookSpecificDateValue( variables.maxDate )
			,"yyyy-MM-dd"
		);
	}

	private void function createIntegerConstraint(){
		if( !Len( variables.maxInteger ) ){
			variables.validationConstraint = variables.dataValidationHelper.createIntegerConstraint(
				getConstraintOperatorObject( "GREATER_OR_EQUAL" )
				,variables.minInteger
				,JavaCast( "null", "" ) //HSSF won't accept an empty string
			);
			return;
		}
		if( !Len( variables.minInteger ) ){
			variables.validationConstraint = variables.dataValidationHelper.createIntegerConstraint(
				getConstraintOperatorObject( "LESS_OR_EQUAL" )
				,variables.maxInteger
				,JavaCast( "null", "" )
			);
			return;
		}
		variables.validationConstraint = variables.dataValidationHelper.createIntegerConstraint(
			getConstraintOperatorObject( "BETWEEN" )
			,variables.minInteger
			,variables.maxInteger
		);
	}

	private void function createListConstraintFromArray(){
		variables.validationConstraint = variables.dataValidationHelper.createExplicitListConstraint( variables.validValues );
	}

	private void function createListConstraintFromCells(){
		var sheetName = determineSourceSheetName();
		var cellReference = variables.library.getRangeHelper().convertRangeReferenceToAbsoluteAddress( variables.valuesSourceCellRange );
		var sheetAndCellReference = quoteSheetNameIfRequired( sheetName ) & "!" & cellReference;
		variables.validationConstraint =  variables.dataValidationHelper.createFormulaListConstraint( sheetAndCellReference );
	}

	private string function getWorkbookSpecificDateValue( required date ){
		if( !IsDate( arguments.date ) ) //formula string
			return arguments.date;
		if( variables.workbookIsXlsx )
			return "Date( #arguments.date.Year()#, #arguments.date.Month()#, #arguments.date.Day()# )";
		return DateFormat( arguments.date, "yyyy-mm-dd" );
	}

	private string function quoteSheetNameIfRequired( required string sheetName ){
		if( arguments.sheetName.REFindNoCase( "\W" ) ) //any non word character: space, hyphen etc... (but not underscore)
			return "'" & sheetName & "'";
		return arguments.sheetName;
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

	private any function getConstraintOperatorObject( required string type ){
		return variables.library.createJavaObject( "org.apache.poi.ss.usermodel.DataValidationConstraint$OperatorType" )[ arguments.type ];
	}

}