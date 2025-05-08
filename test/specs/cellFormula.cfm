<cfscript>
describe( "cellFormula", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( ( wb )=>{
			s.addColumn( wb, "1,1" );
		})
		variables.theFormula = "SUM(A1:A2)";
	})

	it( "Sets and gets the specified formula for the specified cell", ()=>{
		workbooks.Each( ( wb )=>{
			s.setCellFormula( wb, theFormula, 3, 1 );
			expect( s.getCellFormula( wb, 3, 1 ) ).toBe( theFormula );
			expect( s.getCellValue( wb, 3, 1 ) ).toBe( 2 );
		})
	})

	it( "setCellFormula and getCellFormula are chainable", ()=>{
		workbooks.Each( ( wb )=>{
			var actual = s.newChainable( wb )
				.setCellFormula( theFormula, 3, 1 )
				.getCellFormula( 3, 1 );
			expect( actual ).toBe( theFormula );
			expect( s.getCellValue( wb, 3, 1 ) ).toBe( 2 );
		})
	})

	it( "Gets all formulas from the workbook", ()=>{
		workbooks.Each( ( wb )=>{
			s.setCellFormula( wb, theFormula, 3, 1 );
			var expected = [{
				formula: theFormula
				,row: 3
				,column: 1
			}];
			var actual = s.getCellFormula( wb );
			expect( actual ).toBe( expected );
		})
	})

	it( "Returns an empty string if the specified cell doesn't exist", ()=>{
		workbooks.Each( ( wb )=>{
			var actual = s.getCellFormula( wb, 100, 100 );
			expect( actual ).toBeEmpty();
		})
	})

	describe( "recalculation", ()=>{

		it( "can set and get a flag for all formulas in the workbook to be recalculated the next time the file is opened", ()=>{
			workbooks.Each( ( wb )=>{
				expect( s.getRecalculateFormulasOnNextOpen( wb ) ).toBeFalse();
				s.setRecalculateFormulasOnNextOpen( wb );
				expect( s.getRecalculateFormulasOnNextOpen( wb ) ).toBeTrue();
			})
		})

		it( "can set and get a flag for all formulas in a specific sheet to be recalculated the next time the file is opened", ()=>{
			workbooks.Each( ( wb )=>{
				s.createSheet( wb, "sheet2" );
				expect( s.getRecalculateFormulasOnNextOpen( wb, "sheet2" ) ).toBeFalse();
				s.setRecalculateFormulasOnNextOpen( wb, true, "sheet2" );
				expect( s.getRecalculateFormulasOnNextOpen( wb, "sheet2" ) ).toBeTrue();
				expect( s.getRecalculateFormulasOnNextOpen( wb, "sheet1" ) ).toBeFalse();
			})
		})

		it( "setForceFormulaRecalculation on all sheets is chainable", ()=>{
			workbooks.Each( ( wb )=>{
				//all sheets
				var chainable = s.newChainable( wb );
				expect( chainable.getRecalculateFormulasOnNextOpen() ).toBeFalse();
				chainable.setRecalculateFormulasOnNextOpen();
				expect( chainable.getRecalculateFormulasOnNextOpen() ).toBeTrue();
			})
		})

		it( "setForceFormulaRecalculation on a specific sheet is chainable", ()=>{
			workbooks.Each( ( wb )=>{
				chainable = s.newChainable( wb ).createSheet( "sheet2" );
				expect( chainable.getRecalculateFormulasOnNextOpen( "sheet2" ) ).toBeFalse();
				chainable.setRecalculateFormulasOnNextOpen( true, "sheet2" );
				expect( chainable.getRecalculateFormulasOnNextOpen( "sheet2" ) ).toBeTrue();
				expect( chainable.getRecalculateFormulasOnNextOpen( "sheet1" ) ).toBeFalse();
			})
		})

		it( "returns cached calculated values by default", ()=> {
			workbooks.Each( ( wb )=>{
				s.setCellFormula( wb, theFormula, 3, 1 );
				expect( s.getCellValue( wb, 3, 1 ) ).toBe( 2 );
				s.setCellValue( wb, 2, 1, 1 )
				expect( s.getCellValue( wb, 3, 1 ) ).toBe( 2 );
			})
		})

		it( "can return recalculated formula values", ()=> {
			local.s = newSpreadsheetInstance().setReturnCachedFormulaValues( false );
			workbooks.Each( ( wb )=>{
				s.setCellFormula( wb, theFormula, 3, 1 );
				expect( s.getCellValue( wb, 3, 1 ) ).toBe( 2 );
				s.setCellValue( wb, 2, 1, 1 )
				expect( s.getCellValue( wb, 3, 1 ) ).toBe( 3 );
			})
		})

		it( "can force all formulas to be recalculated", ()=> {
			workbooks.Each( ( wb )=>{
				s.setCellFormula( wb, theFormula, 3, 1 );
				expect( s.getCellValue( wb, 3, 1 ) ).toBe( 2 );
				s.setCellValue( wb, 2, 1, 1 )
				expect( s.getCellValue( wb, 3, 1 ) ).toBe( 2 );
				s.recalculateAllFormulas( wb );
				expect( s.getCellValue( wb, 3, 1 ) ).toBe( 3 );
			})
		})
		
		it( "recalculateAllFormulas is chainable", ()=> {
			workbooks.Each( ( wb )=>{
				var chainable = s.newChainable( wb ).setCellFormula( theFormula, 3, 1 );
				expect( chainable.getCellValue( 3, 1 ) ).toBe( 2 );
				chainable.setCellValue( 2, 1, 1 );
				expect( chainable.getCellValue( 3, 1 ) ).toBe( 2 );//cached
				chainable.recalculateAllFormulas();
				expect( chainable.getCellValue( 3, 1 ) ).toBe( 3 );//recalculated
			})
		})

	})

	describe( "evaluation errors", ()=>{

		it( "By default returns the string '##ERROR!' if the formula is malformed", ()=>{
			workbooks.Each( ( wb )=>{
				s.setCellFormula( wb, "SUS(A1:A2)", 3, 1 );
				var actual = s.getCellValue( wb, 3, 1 );
				expect( actual ).toBe( "##ERROR!" );
			})
		})

		it( "By default returns the string '##ERROR!' if the formula evaluates to an error", ()=>{
			workbooks.Each( ( wb )=>{
				s.setCellValue( wb, 0, 2, 1 ).setCellFormula( wb, "A1/A2", 3, 1 );//Divide by zero error
				var actual = s.getCellValue( wb, 3, 1 );
				expect( actual ).toBe( "##ERROR!" );
			})
		})

		it( "Can be configured to throw an exception on any formula evaluation error", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ( wb )=>{
					newSpreadsheetInstance()
						.setThrowExceptionOnFormulaError( true )
						.setCellFormula( wb, "SUS(A1:A2)", 3, 1 )
						.getCellValue( wb, 3, 1 );
				})
				.toThrow( type="cfsimplicity.spreadsheet.failedFormula" );
			})
			workbooks.Each( ( wb )=>{
				expect( ( wb )=>{
					newSpreadsheetInstance()
						.setThrowExceptionOnFormulaError( true )
						.setCellValue( wb, 0, 2, 1 )
						.setCellFormula( wb, "A1/A2", 3, 1 ) //Divide by zero error
						.getCellValue( wb, 3, 1 );
				})
				.toThrow( type="cfsimplicity.spreadsheet.failedFormula" );
			})
		})

	})

})	
</cfscript>