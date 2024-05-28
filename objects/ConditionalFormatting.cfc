component{

	property name="cellRangeReference";
	property name="comparisonOperator";
	property name="format" type="struct";
	property name="formula";
	property name="formula2";
	property name="sheetName";
	property name="sheetNumber";
	/* Internal */
	property name="conditionType" type="string" default="formula";
	property name="library";
	property name="workbookIsXlsx" type="boolean";
	/* POI java objects */
	property name="cellRangeAddress";
	property name="comparisonOperatorObject";
	property name="conditionalFormattingRule";
	property name="sheet";
	property name="sheetConditionalFormatting";
	property name="workbook";

	public ConditionalFormatting function init( required spreadsheetLibrary ){
		variables.library = arguments.spreadsheetLibrary;
		variables.format = {};
		return this;
	}

	/* Public builder API */

	//ACF rejects if()
	public ConditionalFormatting function when( required string formula ){
		variables.formula = arguments.formula;
		return this;
	}

	public ConditionalFormatting function whenCellValueIs( required string comparisonOperator, required string valueOrFormula, string valueOrFormula2 ){
		variables.conditionType = "cellValueIs";
		variables.comparisonOperator = arguments.comparisonOperator;
		variables.formula = arguments.valueOrFormula;
		if( arguments.KeyExists( "valueOrFormula2" ) )
			variables.formula2 = arguments.valueOrFormula2;
		return this;
	}

	public ConditionalFormatting function setFormat( required struct format ){
		variables.format = arguments.format;
		return this;
	}

	public ConditionalFormatting function onCells( required string cellRangeReference ){
		variables.cellRangeReference = arguments.cellRangeReference;
		return this;
	}

	public ConditionalFormatting function onSheetName( required string sheetName ){
		variables.sheetName = arguments.sheetName;
		return this;
	}

	public ConditionalFormatting function onSheetNumber( required string sheetNumber ){
		variables.sheetNumber = arguments.sheetNumber;
		return this;
	}

	public ConditionalFormatting function addToWorkbook( required workbook ){
		variables.workbook = arguments.workbook;
		variables.workbookIsXlsx = variables.library.isXmlFormat( arguments.workbook );
		setSheet();
		setCellRangeAddress();
		variables.sheetConditionalFormatting = variables.sheet.getSheetConditionalFormatting();
		if( variables.conditionType == "cellValueIs" )
			setComparisonRule();
		else
			setFormulaRule();
		setRuleFormat();
		variables.sheetConditionalFormatting.addConditionalFormatting( [ variables.cellRangeAddress ], variables.conditionalFormattingRule );
		return this;
	}

	/* for testing */
	public array function rulesAppliedToCell( required string cellReference ){
		var provider = variables.workbook.getCreationHelper().createFormulaEvaluator();
		var evaluator = variables.library.createJavaObject( "org.apache.poi.ss.formula.ConditionalFormattingEvaluator" ).init( variables.workbook, provider );
		evaluator.clearAllCachedValues();
		var cellReferenceWithSheetName = addSheetNameToCellReference( arguments.cellReference );
		var cellReferenceObject = variables.library.getCellHelper().getReferenceObjectByAddressString( cellReferenceWithSheetName );
		return evaluator.getConditionalFormattingForCell( cellReferenceObject );
	}

	public any function getFirstRuleAppliedToCell( required string cellReference ){
		var rulesApplied = rulesAppliedToCell( arguments.cellReference );
		if( rulesApplied.IsEmpty() )
			return;
		return rulesApplied[ 1 ].getRule();
	}

	public struct function getFormatAppliedToCell( required string cellReference ){
		var result = {};
		var appliedRule = getFirstRuleAppliedToCell( arguments.cellReference );
		if( IsNull( appliedRule ) )
			return result;
		result = buildFontFormatInfo( appliedRule, result );
		result = buildBorderFormatInfo( appliedRule, result );
		return buildPatternFormatInfo( appliedRule, result );
	}

	public ConditionalFormatting function remove( numeric index=0 ){
		variables.sheetConditionalFormatting.removeConditionalFormatting( JavaCast( "int", arguments.index ) );
		return this;
	}

	/* Private */
	private void function setSheet(){
		if( IsNull( variables.sheetName ) && IsNull( variables.sheetNumber ) ){
			variables.sheet = variables.library.getSheetHelper().getActiveSheet( variables.workbook );
			return;
		}
		if( !IsNull( variables.sheetName ) ){
			//name trumps number
			variables.sheet = variables.library.getSheetHelper().getSheetByName( variables.workbook, variables.sheetName );
			return;
		}
		variables.sheet = variables.library.getSheetHelper().getSheetByNumber( variables.workbook, variables.sheetNumber );
	}

	private void function setCellRangeAddress(){
		var cellRangeReferenceWithSheetName = addSheetNameToCellReference( variables.cellRangeReference );
		variables.cellRangeAddress = variables.library.getRangeHelper().getCellRangeAddressFromReference( cellRangeReferenceWithSheetName );
	}

	private void function setComparisonRule(){
		setComparisonOperatorObject();
		if( variables.comparisonOperator.FindNoCase( "BETWEEN" ) ){
			setBetweenRule();
			return;
		}
		variables.conditionalFormattingRule = variables.sheetConditionalFormatting.createConditionalFormattingRule( variables.comparisonOperatorObject, variables.formula );
	}

	private void function setFormulaRule(){
		variables.conditionalFormattingRule = variables.sheetConditionalFormatting.createConditionalFormattingRule( variables.formula );
	}

	private void function setComparisonOperatorObject(){
		var poiOperator = lookupPoiComparisonOperator();
		variables.comparisonOperatorObject = variables.library.createJavaObject( "org.apache.poi.ss.usermodel.ComparisonOperator" )[ poiOperator ];
	}

	private string function lookupPoiComparisonOperator(){
		switch( variables.comparisonOperator ){
			case "GT": case "LT": case "BETWEEN":
				return variables.comparisonOperator;
			case "EQ": return "EQUAL";
			case "NEQ": return "NOT_EQUAL";
			case "GTE": return "GE";
			case "LTE": return "LE";
			case "NOT BETWEEN": return "NOT_BETWEEN";
			default:
				Throw( type=variables.library.getExceptionType() & ".invalidOperatorArgument", message="Invalid comparison operator '#variables.comparisonOperator#'", detail="Valid operators are: EQ, NEQ, GT, LT, GTE, LTE, BETWEEN, and NOT BETWEEN" );
		}
	}

	private void function setBetweenRule(){
		if( IsNull( variables.formula2 ) )
			Throw( type=variables.library.getExceptionType() & ".missingSecondFormulaArgument", message="Missing formula2 argument", detail="BETWEEN comparisons require two values or formulas" );
		variables.conditionalFormattingRule = variables.sheetConditionalFormatting.createConditionalFormattingRule( variables.comparisonOperatorObject, variables.formula, variables.formula2 );
	}

	private void function setRuleFormat(){
		for( var setting in variables.format ){
			setFormatFromSetting( setting );
		}
	}

	private string function addSheetNameToCellReference( required string cellReference ){
		return variables.sheet.getSheetName() & "!" & arguments.cellReference;
	}

	private struct function buildBorderFormatInfo( required rule, required struct info ){
		if( IsNull( rule.getBorderFormatting() ) )
			return arguments.info;
		var result = arguments.info;
		result.bottomBorderColor = getRGBfromColorObject( arguments.rule.getBorderFormatting().getBottomBorderColorColor() );
		result.bottomBorder = arguments.rule.getBorderFormatting().getBorderBottom().toString();
		result.leftBorderColor = getRGBfromColorObject( arguments.rule.getBorderFormatting().getLeftBorderColorColor() );
		result.leftBorder = arguments.rule.getBorderFormatting().getBorderLeft().toString();
		result.rightBorderColor = getRGBfromColorObject( arguments.rule.getBorderFormatting().getRightBorderColorColor() );
		result.rightBorder = arguments.rule.getBorderFormatting().getBorderRight().toString();
		result.topBorderColor = getRGBfromColorObject( arguments.rule.getBorderFormatting().getTopBorderColorColor() );
		result.topBorder = arguments.rule.getBorderFormatting().getBorderTop().toString();
		return result;
	}

	private struct function buildFontFormatInfo( required rule, required struct info ){
		if( IsNull( rule.getFontFormatting() ) )
			return arguments.info;
		var result = arguments.info;
		result.fontColor = getRGBfromColorObject( arguments.rule.getFontFormatting().getFontColor() );
		var twips = arguments.rule.getFontFormatting().getFontHeight();
		var points = ( twips / 20 );
		result.fontSize = points;
		result.bold = arguments.rule.getFontFormatting().isBold();
		result.italic = arguments.rule.getFontFormatting().isItalic();
		result.underline = variables.library.getFormatHelper().underlineNameFromIndex( arguments.rule.getFontFormatting().getUnderlineType() );
		return result;
	}

	private struct function buildPatternFormatInfo( required rule, required struct info ){
		if( IsNull( rule.getPatternFormatting() ) )
			return arguments.info;
		var result = arguments.info;
		result.backgroundFillColor = getRGBfromColorObject( arguments.rule.getPatternFormatting().getFillBackgroundColorColor() );
		result.foregroundFillColor = getRGBfromColorObject( arguments.rule.getPatternFormatting().getFillForegroundColorColor() );
		result.fillPattern = variables.library.getFormatHelper().patternNameFromIndex( arguments.rule.getPatternFormatting().getFillPattern() );
		return result;
	}

	private void function setFormatFromSetting( required string setting ){
		var settingValue = variables.format[ arguments.setting ];
		if( setting.FindNoCase( "border" ) )
			ensureBorderFormattingExists();
		if( setting.FindNoCase( "fill" ) )
			ensurePatternFormattingExists();
		switch( arguments.setting ){
			case "fontColor":
				ensureFontFormattingExists();
				var colorValue = getColorObjectOrIndex( settingValue );
				if( IsNumeric( colorValue ) )
					variables.conditionalFormattingRule.getFontFormatting().setFontColorIndex( JavaCast( "short", colorValue ) );
				else
					variables.conditionalFormattingRule.getFontFormatting().setFontColor( colorValue );
				return;
			case "fontSize":
				ensureFontFormattingExists();
				var twips = ( settingValue * 20 ); //points to twips
				variables.conditionalFormattingRule.getFontFormatting().setFontHeight( JavaCast( "int", twips ) );
				return;
			case "bold": case "italic":
				ensureFontFormattingExists();
				var boldSetting = variables.format.bold?:false;
				var italicSetting = variables.format.italic?:false;
				variables.conditionalFormattingRule.getFontFormatting().setFontStyle( JavaCast( "boolean", italicSetting ), JavaCast( "boolean", boldSetting ) );
				return;
			case "underline":
				ensureFontFormattingExists();
				var underlineType = underlineIndexFromValue( settingValue );
				if( underlineType == -1 )
					return;
				variables.conditionalFormattingRule.getFontFormatting().setUnderlineType( JavaCast( "byte", underlineType ) );
				return;
			case "bottomBorder":
				var borderStyle = variables.conditionalFormattingRule.getBorderFormatting().getBorderBottom()[ UCase( settingValue ) ];
				variables.conditionalFormattingRule.getBorderFormatting().setBorderBottom( borderStyle );
				return;
			case "bottomBorderColor":
				variables.conditionalFormattingRule.getBorderFormatting().setBottomBorderColor( getColorObjectOrIndex( settingValue ) );
				return;
			case "leftBorder":
				var borderStyle = variables.conditionalFormattingRule.getBorderFormatting().getBorderLeft()[ UCase( settingValue ) ];
				variables.conditionalFormattingRule.getBorderFormatting().setBorderLeft( borderStyle );
				return;
			case "leftBorderColor":
				variables.conditionalFormattingRule.getBorderFormatting().setLeftBorderColor( getColorObjectOrIndex( settingValue ) );
				return;
			case "rightBorder":
				var borderStyle = variables.conditionalFormattingRule.getBorderFormatting().getBorderRight()[ UCase( settingValue ) ];
				variables.conditionalFormattingRule.getBorderFormatting().setBorderRight( borderStyle );
				return;
			case "rightBorderColor":
				variables.conditionalFormattingRule.getBorderFormatting().setRightBorderColor( getColorObjectOrIndex( settingValue ) );
				return;
			case "topBorder":
				var borderStyle = variables.conditionalFormattingRule.getBorderFormatting().getBorderTop()[ UCase( settingValue ) ];
				variables.conditionalFormattingRule.getBorderFormatting().setBorderTop( borderStyle );
				return;
			case "topBorderColor":
				variables.conditionalFormattingRule.getBorderFormatting().setTopBorderColor( getColorObjectOrIndex( settingValue ) );
				return;
			//NOTE: fillColors appear to be limited in what they will accept: basic colors only it seems
			case "backgroundFillColor":
				variables.conditionalFormattingRule.getPatternFormatting().setFillBackgroundColor( getColorObjectOrIndex( settingValue ) );
				return;
			case "foregroundFillColor":
				variables.conditionalFormattingRule.getPatternFormatting().setFillForegroundColor( getColorObjectOrIndex( settingValue ) );
				return;
			case "fillPattern":
				var fillPattern = variables.conditionalFormattingRule.getPatternFormatting()[ UCase( settingValue ) ];
				variables.conditionalFormattingRule.getPatternFormatting().setFillPattern( fillPattern );
				return;
		}
	}

	private void function ensureFontFormattingExists(){
		if( IsNull( variables.conditionalFormattingRule.getFontFormatting() ) )
			variables.conditionalFormattingRule.createFontFormatting();
	}

	private void function ensureBorderFormattingExists(){
		if( IsNull( variables.conditionalFormattingRule.getBorderFormatting() ) )
			variables.conditionalFormattingRule.createBorderFormatting();
	}

	private void function ensurePatternFormattingExists(){
		if( IsNull( variables.conditionalFormattingRule.getPatternFormatting() ) )
			variables.conditionalFormattingRule.createPatternFormatting();
	}

	private any function getColorObjectOrIndex( required string settingValue ){
		return variables.library.getColorHelper().getColor( variables.workbook, arguments.settingValue );
	}

	private any function getRGBfromColorObject( required any colorObject ){
		return variables.library.getColorHelper().getRGBStringFromColorObject( arguments.colorObject );
	}

	private numeric function underlineIndexFromValue( required any value ){
		switch( arguments.value ){
			case "none": return 0;
			case "single": return 1;
			case "double": return 2;
			//NB: accounting underlines not supported
		}
		if( IsBoolean( arguments.value ) )
			return arguments.value? 1: 0;
		return -1;
	}

}