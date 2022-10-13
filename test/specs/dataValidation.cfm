<cfscript>
describe( "dataValidation", function(){

	beforeEach( function(){
		variables.cellRange = "A1:B1";
		variables.validValues = [ "London", "Paris", "New York" ];
	});

	it( "can create a validation drop-down using an array of values", function() {
		var dv = s.newDataValidation()
			.onCells( cellRange )
			.withValues( validValues );
		variables.spreadsheetTypes.Each( function( type ){
			var chainable = s.newChainable( type ).addDataValidation( dv );
			expect( dv.validValueArrayAppliedToSheet() ).toBe( validValues );
			expect( dv.targetCellRangeAppliedToSheet() ).toBe( cellRange );
		});
		//alternate direct syntax
		variables.spreadsheetTypes.Each( function( type ){
			var wb = ( type == "xls" )? s.newXls(): s.newXlsx();
			var dv = s.newDataValidation()
				.onCells( cellRange )
				.withValues( validValues )
				.addToWorkbook( wb );
			expect( dv.targetCellRangeAppliedToSheet() ).toBe( cellRange );
			expect( dv.validValueArrayAppliedToSheet() ).toBe( validValues );
		});
	});

	it( "can create a validation drop-down from values in other cells in the same sheet", function() {
		var dv = s.newDataValidation()
			.onCells( cellRange )
			.withValuesFromCells( "C1:C3" );
		variables.spreadsheetTypes.Each( function( type ){
			var chainable = s.newChainable( type )
				.addColumn( data=validValues, startColumn=3 )
				.addDataValidation( dv );
			expect( dv.targetCellRangeAppliedToSheet() ).toBe( cellRange );
			expect( dv.sourceCellsReferenceAppliedToSheet() ).toBe( "Sheet1!$C$1:$C$3" );
		});
	});

	it( "can create a validation drop-down from values in other cells in a different sheet", function() {
		var dv = s.newDataValidation()
			.onCells( cellRange )
			.withValuesFromSheetName( "cities" )
			.withValuesFromCells( "A1:A3" );
		variables.spreadsheetTypes.Each( function( type ){
			var chainable = s.newChainable( type )
				.createSheet( "cities" )
				.setActiveSheetNumber( 2 )
				.addColumn( data=validValues )
				.setActiveSheetNumber( 1 )
				.addDataValidation( dv );
			expect( dv.targetCellRangeAppliedToSheet() ).toBe( cellRange );
			expect( dv.sourceCellsReferenceAppliedToSheet() ).toBe( "cities!$A$1:$A$3" );
		});
	});

	it( "can create a validation drop-down from values in other cells in a different sheet the name of which includes a space", function() {
		var dv = s.newDataValidation()
			.onCells( cellRange )
			.withValuesFromSheetName( "towns and cities" )
			.withValuesFromCells( "A1:A3" );
		variables.spreadsheetTypes.Each( function( type ){
			var chainable = s.newChainable( type )
				.createSheet( "towns and cities" )
				.setActiveSheetNumber( 2 )
				.addColumn( data=validValues )
				.setActiveSheetNumber( 1 )
				.addDataValidation( dv );
			expect( dv.targetCellRangeAppliedToSheet() ).toBe( cellRange );
			expect( dv.sourceCellsReferenceAppliedToSheet() ).toBe( "'towns and cities'!$A$1:$A$3" );
		});
	});

	it( "allows the validation error message to be customised", function() {
		var errorTitle = "Wrong";
		var errorMessage = "Think again, dude.";
		var dv = s.newDataValidation()
			.onCells( cellRange )
			.withValues( validValues )
			.withErrorTitle( errorTitle )
			.withErrorMessage( errorMessage );
		variables.spreadsheetTypes.Each( function( type ){
			var chainable = s.newChainable( type ).addDataValidation( dv );
			expect( dv.errorTitleAppliedToSheet() ).toBe( errorTitle );
			expect( dv.errorMessageAppliedToSheet() ).toBe( errorMessage );
		});
	});

	it( "the drop-down arrow can be suppressed for a passed array of data", function() {
		
		var dv = s.newDataValidation()
			.onCells( cellRange )
			.withValues( validValues )
			.withNoDropdownArrow();
		variables.spreadsheetTypes.Each( function( type ){
			var chainable = s.newChainable( type ).addDataValidation( dv );
			var falseForXlsxTrueForXls = ( type != "xlsx" );// XLSX requires the OPPOSITE explicit boolean setting (WTF!)
			expect( dv.suppressDropdownSettingArrowAppliedToSheet() ).toBe( falseForXlsxTrueForXls );
		});
	});

	describe( "throws an exception if", function(){

		it( "the specified source sheet doesn't exist", function(){
			var dv = s.newDataValidation()
				.onCells( cellRange )
				.withValuesFromSheetName( "nonexistant" )
				.withValuesFromCells( "A1:A3" );
			variables.spreadsheetTypes.Each( function( type ){
				expect( function(){
					var chainable = s.newChainable( type ).addDataValidation( dv );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetName" );
			});

		});

	});


});	
</cfscript>