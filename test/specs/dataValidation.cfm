<cfscript>
describe( "dataValidation", ()=>{

	beforeEach( ()=>{
		variables.cellRange = "A1:B1";
		variables.validValues = [ "London", "Paris", "New York" ];
		variables.minDate = CreateDate( 2020, 1, 1 );
		variables.maxDate = CreateDate( 2020, 12, 31 );
	})

	describe( "drop-downs", ()=>{

		it( "can create a validation drop-down using an array of values", ()=>{
			variables.spreadsheetTypes.Each( ( type )=>{
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withValues( validValues );
				var chainable = s.newChainable( type ).addDataValidation( dv );
				expect( dv.validValueArrayAppliedToSheet() ).toBe( validValues );
				expect( dv.targetCellRangeAppliedToSheet() ).toBe( cellRange );
			})
			//alternate direct syntax
			variables.spreadsheetTypes.Each( ( type )=>{
				var wb = ( type == "xls" )? s.newXls(): s.newXlsx();
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withValues( validValues )
					.addToWorkbook( wb );
				expect( dv.targetCellRangeAppliedToSheet() ).toBe( cellRange );
				expect( dv.validValueArrayAppliedToSheet() ).toBe( validValues );
			})
		})

		it( "can create a validation drop-down from values in other cells in the same sheet", ()=>{
			variables.spreadsheetTypes.Each( ( type )=>{
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withValuesFromCells( "C1:C3" );
				var chainable = s.newChainable( type )
					.addColumn( data=validValues, startColumn=3 )
					.addDataValidation( dv );
				expect( dv.targetCellRangeAppliedToSheet() ).toBe( cellRange );
				expect( dv.sourceCellsReferenceAppliedToSheet() ).toBe( "Sheet1!$C$1:$C$3" );
			})
		})

		it( "can create a validation drop-down from values in other cells in a different sheet", ()=>{
			variables.spreadsheetTypes.Each( ( type )=>{
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withValuesFromSheetName( "cities" )
					.withValuesFromCells( "A1:A3" );
				var chainable = s.newChainable( type )
					.createSheet( "cities" )
					.setActiveSheetNumber( 2 )
					.addColumn( data=validValues )
					.setActiveSheetNumber( 1 )
					.addDataValidation( dv );
				expect( dv.targetCellRangeAppliedToSheet() ).toBe( cellRange );
				expect( dv.sourceCellsReferenceAppliedToSheet() ).toBe( "cities!$A$1:$A$3" );
			})
		})

		it( "can create a validation drop-down from values in other cells in a different sheet the name of which includes a space", ()=>{
			variables.spreadsheetTypes.Each( ( type )=>{
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withValuesFromSheetName( "towns and cities" )
					.withValuesFromCells( "A1:A3" );
				var chainable = s.newChainable( type )
					.createSheet( "towns and cities" )
					.setActiveSheetNumber( 2 )
					.addColumn( data=validValues )
					.setActiveSheetNumber( 1 )
					.addDataValidation( dv );
				expect( dv.targetCellRangeAppliedToSheet() ).toBe( cellRange );
				expect( dv.sourceCellsReferenceAppliedToSheet() ).toBe( "'towns and cities'!$A$1:$A$3" );
			})
		})

		it( "the drop-down arrow can be suppressed for a passed array of data", ()=>{
			variables.spreadsheetTypes.Each( ( type )=>{
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withValues( validValues )
					.withNoDropdownArrow();
				var chainable = s.newChainable( type ).addDataValidation( dv );
				var falseForXlsxTrueForXls = ( type != "xlsx" );// XLSX requires the OPPOSITE explicit boolean setting (WTF!)
				expect( dv.suppressDropdownSettingArrowAppliedToSheet() ).toBe( falseForXlsxTrueForXls );
			})
		})

	})

	describe( "date constraints", ()=>{

		it( "can constrain input to a minimum date", ()=>{
			variables.spreadsheetTypes.Each( ( type )=>{
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withMinDate( minDate );
				var chainable = s.newChainable( type ).addDataValidation( dv );
				expect( dv.getConstraintType() ).toBe( "date" );
				expect( dv.getConstraintOperator() ).toBe( "GREATER_OR_EQUAL" );
			})
		})

		it( "can constrain input to a maximum date", ()=>{
			variables.spreadsheetTypes.Each( ( type )=>{
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withMaxDate( minDate );
				var chainable = s.newChainable( type ).addDataValidation( dv );
				expect( dv.getConstraintType() ).toBe( "date" );
				expect( dv.getConstraintOperator() ).toBe( "LESS_OR_EQUAL" );
			})
		})

		it( "can constrain input to a date range", ()=>{
			variables.spreadsheetTypes.Each( ( type )=>{
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withMinDate( minDate )
					.withMaxDate( maxDate );
				var chainable = s.newChainable( type ).addDataValidation( dv );
				expect( dv.getConstraintType() ).toBe( "date" );
				expect( dv.getConstraintOperator() ).toBe( "BETWEEN" );
			})
		})

		it( "can constrain input to a dates derived from formulas", ()=>{
			variables.spreadsheetTypes.Each( ( type )=>{
				var dv = s.newDataValidation()
					.onCells( "B1" )
					.withMinDate( "=$A$1" )
					.withMaxDate( "=$A$2" );
				var chainable = s.newChainable( type )
					.addColumn( [ minDate, maxDate ] ) //A1-2
					.addDataValidation( dv );
				expect( dv.getConstraintType() ).toBe( "date" );
				expect( dv.getConstraintOperator() ).toBe( "BETWEEN" );
			})
		})

	})

	describe( "integer constraints", ()=>{

		it("can constrain input to a minimum integer", () => {
			variables.spreadsheetTypes.Each( ( type ) => {
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withMinInteger( 140 );
				var chainable = s.newChainable( type ).addDataValidation( dv );
				expect( dv.getConstraintType() ).toBe( "integer" );
				expect( dv.getConstraintOperator() ).toBe( "GREATER_OR_EQUAL" );
			})
		})

		it("can constrain input to a maximum integer", () => {
			variables.spreadsheetTypes.Each( ( type ) => {
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withMaxInteger( 5 );
				var chainable = s.newChainable( type ).addDataValidation( dv );
				expect( dv.getConstraintType() ).toBe( "integer" );
				expect( dv.getConstraintOperator() ).toBe( "LESS_OR_EQUAL" );
			})
		})

		it( "can constrain input to an integer range", ()=>{
			variables.spreadsheetTypes.Each( ( type )=>{
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withMinInteger( 0 )
					.withMaxInteger( 100 );
				var chainable = s.newChainable( type ).addDataValidation( dv );
				expect( dv.getConstraintType() ).toBe( "integer" );
				expect( dv.getConstraintOperator() ).toBe( "BETWEEN" );
			})
		})

		it( "can constrain input to integers derived from formulas", ()=>{
			variables.spreadsheetTypes.Each( ( type )=>{
				var dv = s.newDataValidation()
					.onCells( "C1" )
					.withMinInteger( "=SUM($A$1:$A$3)" )
					.withMaxInteger( "=SUM($B$1:$B$3)" );
				var chainable = s.newChainable( type )
					.addColumn( [ 1, 1, 1 ] ) //A1-3
					.addColumn( [ 2, 2, 2 ] ) //B1-3
					.addDataValidation( dv );
				expect( dv.getConstraintType() ).toBe( "integer" );
				expect( dv.getConstraintOperator() ).toBe( "BETWEEN" );
			})
		})

	})

	describe( "decimal constraints", ()=>{

		it("can constrain input to a minimum decimal", () => {
			variables.spreadsheetTypes.Each( ( type ) => {
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withMinDecimal( 1.5 );
				var chainable = s.newChainable( type ).addDataValidation( dv );
				expect( dv.getConstraintType() ).toBe( "decimal" );
				expect( dv.getConstraintOperator() ).toBe( "GREATER_OR_EQUAL" );
			})
		})

		it("can constrain input to a maximum decimal", () => {
			variables.spreadsheetTypes.Each( ( type ) => {
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withMaxDecimal( 2.5 );
				var chainable = s.newChainable( type ).addDataValidation( dv );
				expect( dv.getConstraintType() ).toBe( "decimal" );
				expect( dv.getConstraintOperator() ).toBe( "LESS_OR_EQUAL" );
			})
		})

		it( "can constrain input to a decimal range", ()=>{
			variables.spreadsheetTypes.Each( ( type )=>{
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withMinDecimal( 1.5 )
					.withMaxDecimal( 2.5 );
				var chainable = s.newChainable( type ).addDataValidation( dv );
				expect( dv.getConstraintType() ).toBe( "decimal" );
				expect( dv.getConstraintOperator() ).toBe( "BETWEEN" );
			})
		})

		it( "can constrain input to decimals derived from formulas", ()=>{
			variables.spreadsheetTypes.Each( ( type )=>{
				var dv = s.newDataValidation()
					.onCells( "C1" )
					.withMinDecimal( "=SUM($A$1:$A$3)" )
					.withMaxDecimal( "=SUM($B$1:$B$3)" );
				var chainable = s.newChainable( type )
					.addColumn( [ 1.1, 1.1, 1.1 ] ) //A1-3
					.addColumn( [ 2.2, 2.2, 2.2 ] ) //B1-3
					.addDataValidation( dv );
				expect( dv.getConstraintType() ).toBe( "decimal" );
				expect( dv.getConstraintOperator() ).toBe( "BETWEEN" );
			})
		})

	})

	describe( "text length constraints", ()=>{

		it("can constrain input to a minimum text length", () => {
			variables.spreadsheetTypes.Each( ( type ) => {
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withMinLength( 5 );
				var chainable = s.newChainable( type ).addDataValidation( dv );
				expect( dv.getConstraintType() ).toBe( "length" );
				expect( dv.getConstraintOperator() ).toBe( "GREATER_OR_EQUAL" );
			})
		})

		it("can constrain input to a maximum text length", () => {
			variables.spreadsheetTypes.Each( ( type ) => {
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withMaxLength( 10 );
				var chainable = s.newChainable( type ).addDataValidation( dv );
				expect( dv.getConstraintType() ).toBe( "length" );
				expect( dv.getConstraintOperator() ).toBe( "LESS_OR_EQUAL" );
			})
		})

		it( "can constrain input to a text length range", ()=>{
			variables.spreadsheetTypes.Each( ( type )=>{
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withMinLength( 5 )
					.withMaxLength( 10 );
				var chainable = s.newChainable( type ).addDataValidation( dv );
				expect( dv.getConstraintType() ).toBe( "length" );
				expect( dv.getConstraintOperator() ).toBe( "BETWEEN" );
			})
		})

	})

	describe( "custom constraints", ()=>{

		it( "can constrain using a formula which must evaluate to true", ()=>{
			variables.spreadsheetTypes.Each( ( type ) => {
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withFormula( "ISTEXT(A1)" );
				var chainable = s.newChainable( type ).addDataValidation( dv );
				expect( dv.getConstraintType() ).toBe( "formula" );
			})
		})

		it( "will strip any leading '=' from the formula which POI won't allow", ()=>{
			variables.spreadsheetTypes.Each( ( type ) => {
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withFormula( "=ISTEXT(A1)" );
				var chainable = s.newChainable( type ).addDataValidation( dv );
				expect( dv.getConstraintType() ).toBe( "formula" );
			})
		})

	})

	it( "knows its constraint type", ()=>{
		variables.spreadsheetTypes.Each( ( type )=>{
			var dv = s.newDataValidation()
				.onCells( cellRange )
				.withValues( validValues );
			var chainable = s.newChainable( type ).addDataValidation( dv );
			expect( dv.getConstraintType() ).toBe( "list" );
		})
	})

	it( "allows the validation error message to be customised", ()=>{
		var errorTitle = "Wrong";
		var errorMessage = "Think again!";
		variables.spreadsheetTypes.Each( ( type )=>{
			var dv = s.newDataValidation()
				.onCells( cellRange )
				.withValues( validValues )
				.withErrorTitle( errorTitle )
				.withErrorMessage( errorMessage );
			var chainable = s.newChainable( type ).addDataValidation( dv );
			expect( dv.errorTitleAppliedToSheet() ).toBe( errorTitle );
			expect( dv.errorMessageAppliedToSheet() ).toBe( errorMessage );
		})
	})

	describe( "throws an exception if", ()=>{

		it( "the specified source sheet doesn't exist", ()=>{
			variables.spreadsheetTypes.Each( ( type )=>{
				var dv = s.newDataValidation()
					.onCells( cellRange )
					.withValuesFromSheetName( "nonexistant" )
					.withValuesFromCells( "A1:A3" );
				expect( ()=>{
					var chainable = s.newChainable( type ).addDataValidation( dv );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetName" );
			})

		})

	})

})	
</cfscript>