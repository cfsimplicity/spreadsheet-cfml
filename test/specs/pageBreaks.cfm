<cfscript>
	describe( "pageBreaks", ()=>{

		beforeEach( ()=>{
			variables.workbooks = [ s.newXls(), s.newXlsx() ];
			var columnData = [ "a", "b", "c", "d", "e" ];
			workbooks.Each( ( wb )=>{
				s.addRows( wb, [ columnData, columnData, columnData, columnData, columnData ] );
			})
		})

		it( "can be inserted and removed after a specific row", ()=> {
			workbooks.Each( ( wb )=>{
				s.setRowBreak( wb, 3 );
				expect( s.getSheetHelper().getActiveSheet( wb ).IsRowBroken( 2 ) ).toBeTrue();
				s.removeRowBreak( wb, 3 );
				expect( s.getSheetHelper().getActiveSheet( wb ).IsRowBroken( 2 ) ).toBeFalse();
			})
			workbooks.Each( ( wb )=>{
				var chainable = s.newChainable( wb ).setRowBreak( 3 );
				expect( s.getSheetHelper().getActiveSheet( wb ).IsRowBroken( 2 ) ).toBeTrue();
				chainable.removeRowBreak( 3 );
				expect( s.getSheetHelper().getActiveSheet( wb ).IsRowBroken( 2 ) ).toBeFalse();
			})
		})

		it( "can be inserted and removed after a specific column", ()=> {
			workbooks.Each( ( wb )=>{
				s.setColumnBreak( wb, 3 );
				expect( s.getSheetHelper().getActiveSheet( wb ).IsColumnBroken( 2 ) ).toBeTrue();
				s.removeColumnBreak( wb, 3 );
				expect( s.getSheetHelper().getActiveSheet( wb ).IsColumnBroken( 2 ) ).toBeFalse();
			})
			workbooks.Each( ( wb )=>{
				var chainable = s.newChainable( wb ).setColumnBreak( 3 );
				expect( s.getSheetHelper().getActiveSheet( wb ).IsColumnBroken( 2 ) ).toBeTrue();
				chainable.removeColumnBreak( 3 );
				expect( s.getSheetHelper().getActiveSheet( wb ).IsColumnBroken( 2 ) ).toBeFalse();
			})
		})
	
		describe( "addPageBreaks allows multiple breaks to be inserted", ()=>{

			it( "adds page breaks at the row and column numbers passed in as lists", ()=>{
				workbooks.Each( ( wb )=>{
					s.addPageBreaks( wb, "2,3", "1,2" );
					expect( s.getSheetHelper().getActiveSheet( wb ).getRowBreaks() ).toBe( [ 1, 2 ] );
					expect( s.getSheetHelper().getActiveSheet( wb ).getColumnBreaks() ).toBe( [ 0, 1 ] );
				})
			})

			it( "Doesn't error when passing valid arguments with extra trailing/leading space", ()=>{
				workbooks.Each( ( wb )=>{
					s.addPageBreaks( wb, " 2,3 ", "1,2 " );
				})
			})

			it( "Doesn't error when passing single numbers instead of lists", ()=>{
				workbooks.Each( ( wb )=>{
					s.addPageBreaks( wb, 1, 2 );
				})
			})

			it( "Is chainable", ()=>{
				workbooks.Each( ( wb )=>{
					s.newChainable( wb ).addPageBreaks( 1, 2 );
				})
			})

			it( "Throws a helpful exception if both arguments are missing or present but empty", ()=>{
				workbooks.Each( ( wb )=>{
					expect( ()=>{
						s.addPageBreaks( wb );
					}).toThrow( type="cfsimplicity.spreadsheet.missingRowOrColumnBreaksArgument" );
					expect( ()=>{
						s.addPageBreaks( wb, "" );
					}).toThrow( type="cfsimplicity.spreadsheet.missingRowOrColumnBreaksArgument" );
				})
			})

		})

	})
</cfscript>