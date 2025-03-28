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

		it( "can set a flag for all formulas to be recalculated in the workbook the next time the file is opened", ()=>{
			// only xlsx seems to reflect the flag status
			var wb = s.newXlsx();
			expect( wb.getForceFormulaRecalculation() ).toBeFalse();
			s.setRecalculateFormulasOnNextOpen( wb );
			expect( wb.getForceFormulaRecalculation() ).toBeTrue();
		})

		it( "getForceFormulaRecalculation is chainable", ()=>{
			// only xlsx seems to reflect the flag status
			var wb = s.newXlsx();
			expect( wb.getForceFormulaRecalculation() ).toBeFalse();
			s.newChainable( wb ).setRecalculateFormulasOnNextOpen();
			expect( wb.getForceFormulaRecalculation() ).toBeTrue();
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

})	
</cfscript>