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
			var dv = s.newDataValidation()
				.onCells( cellRange )
				.withValues( validValues );
			variables.spreadsheetTypes.Each( ( type )=>{
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
			var dv = s.newDataValidation()
				.onCells( cellRange )
				.withValuesFromCells( "C1:C3" );
			variables.spreadsheetTypes.Each( ( type )=>{
				var chainable = s.newChainable( type )
					.addColumn( data=validValues, startColumn=3 )
					.addDataValidation( dv );
				expect( dv.targetCellRangeAppliedToSheet() ).toBe( cellRange );
				expect( dv.sourceCellsReferenceAppliedToSheet() ).toBe( "Sheet1!$C$1:$C$3" );
			})
		})

		it( "can create a validation drop-down from values in other cells in a different sheet", ()=>{
			var dv = s.newDataValidation()
				.onCells( cellRange )
				.withValuesFromSheetName( "cities" )
				.withValuesFromCells( "A1:A3" );
			variables.spreadsheetTypes.Each( ( type )=>{
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
			var dv = s.newDataValidation()
				.onCells( cellRange )
				.withValuesFromSheetName( "towns and cities" )
				.withValuesFromCells( "A1:A3" );
			variables.spreadsheetTypes.Each( ( type )=>{
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
			var dv = s.newDataValidation()
				.onCells( cellRange )
				.withValues( validValues )
				.withNoDropdownArrow();
			variables.spreadsheetTypes.Each( ( type )=>{
				var chainable = s.newChainable( type ).addDataValidation( dv );
				var falseForXlsxTrueForXls = ( type != "xlsx" );// XLSX requires the OPPOSITE explicit boolean setting (WTF!)
				expect( dv.suppressDropdownSettingArrowAppliedToSheet() ).toBe( falseForXlsxTrueForXls );
			})
		})

	})

	describe( "date constraints", ()=>{

		it( "can constrain input to a date range", ()=>{
			var dv = s.newDataValidation()
				.onCells( cellRange )
				.withMinDate( minDate )
				.withMaxDate( maxDate );
			variables.spreadsheetTypes.Each( ( type )=>{
				var chainable = s.newChainable( type ).addDataValidation( dv );
				expect( dv.getConstraintType() ).toBe( "date" );
				expect( dv.getConstraintOperator() ).toBe( "BETWEEN" );
			})
		})

	})

	describe( "integer constraints", ()=>{

		it( "can constrain input to an integer range", ()=>{
			var dv = s.newDataValidation()
				.onCells( cellRange )
				.withMinInteger( 0 )
				.withMaxInteger( 100 );
			variables.spreadsheetTypes.Each( ( type )=>{
				var chainable = s.newChainable( type ).addDataValidation( dv );
				expect( dv.getConstraintType() ).toBe( "integer" );
				expect( dv.getConstraintOperator() ).toBe( "BETWEEN" );
			})
		})

	})

	it( "knows its constraint type", ()=>{
		var dv = s.newDataValidation()
			.onCells( cellRange )
			.withValues( validValues );
		variables.spreadsheetTypes.Each( ( type )=>{
			var chainable = s.newChainable( type ).addDataValidation( dv );
			expect( dv.getConstraintType() ).toBe( "list" );
		})
	})

	it( "allows the validation error message to be customised", ()=>{
		var errorTitle = "Wrong";
		var errorMessage = "Think again, dude.";
		var dv = s.newDataValidation()
			.onCells( cellRange )
			.withValues( validValues )
			.withErrorTitle( errorTitle )
			.withErrorMessage( errorMessage );
		variables.spreadsheetTypes.Each( ( type )=>{
			var chainable = s.newChainable( type ).addDataValidation( dv );
			expect( dv.errorTitleAppliedToSheet() ).toBe( errorTitle );
			expect( dv.errorMessageAppliedToSheet() ).toBe( errorMessage );
		})
	})

	describe( "throws an exception if", ()=>{

		it( "the specified source sheet doesn't exist", ()=>{
			var dv = s.newDataValidation()
				.onCells( cellRange )
				.withValuesFromSheetName( "nonexistant" )
				.withValuesFromCells( "A1:A3" );
			variables.spreadsheetTypes.Each( ( type )=>{
				expect( ()=>{
					var chainable = s.newChainable( type ).addDataValidation( dv );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetName" );
			})

		})

		it( "a minDate is specified but no maxDate or vice versa", ()=>{
			var dv = s.newDataValidation()
				.onCells( cellRange )
				.withMinDate( minDate );
			variables.spreadsheetTypes.Each( ( type )=>{
				expect( ()=>{
					var chainable = s.newChainable( type ).addDataValidation( dv );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidValidationConstraint" );
			})
			dv = s.newDataValidation()
				.onCells( cellRange )
				.withMaxDate( maxDate );
			variables.spreadsheetTypes.Each( ( type )=>{
				expect( ()=>{
					var chainable = s.newChainable( type ).addDataValidation( dv );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidValidationConstraint" );
			})

		})

		it( "a minInteger is specified but no maxInteger or vice versa", ()=>{
			var dv = s.newDataValidation()
				.onCells( cellRange )
				.withMinInteger( 0 );
			variables.spreadsheetTypes.Each( ( type )=>{
				expect( ()=>{
					var chainable = s.newChainable( type ).addDataValidation( dv );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidValidationConstraint" );
			})
			dv = s.newDataValidation()
				.onCells( cellRange )
				.withMaxInteger( 100 );
			variables.spreadsheetTypes.Each( ( type )=>{
				expect( ()=>{
					var chainable = s.newChainable( type ).addDataValidation( dv );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidValidationConstraint" );
			})

		})

	})

})	
</cfscript>