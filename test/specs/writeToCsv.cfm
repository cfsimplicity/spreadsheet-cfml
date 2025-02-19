<cfscript>
describe( "writeToCsv", ()=>{

	beforeEach( ()=>{
		var data = [ [ "a", "b" ], [ "c", "d" ] ];
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( ( wb )=>{
			s.addRows( wb, data );
		})
	})

	it( "writes a csv file from a spreadsheet object", ()=>{
		var expectedCsv = 'a,b#newline#c,d#newline#';
		workbooks.Each( ( wb )=>{
			s.writeToCsv( wb, tempCsvPath, true );
			expect( FileRead( tempCsvPath ) ).toBe( expectedCsv );
		})
	})

	it( "is chainable", ()=>{
		var expectedCsv = 'a,b#newline#c,d#newline#';
		workbooks.Each( ( wb )=>{
			s.newChainable( wb ).writeToCsv( tempCsvPath, true );
			expect( FileRead( tempCsvPath ) ).toBe( expectedCsv );
		})
	})

	it( "allows an alternative delimiter", ()=>{
		var expectedCsv = 'a|b#newline#c|d#newline#';
		workbooks.Each( ( wb )=>{
			s.writeToCsv( wb, tempCsvPath, true, "|" );
			expect( FileRead( tempCsvPath ) ).toBe( expectedCsv );
		})
	})

	it( "allows the sheet's header row to be excluded", ()=>{
		var expectedCsv = 'a,b#newline#c,d#newline#';
		workbooks.Each( ( wb )=>{
			s.addRow( wb, [ "column1", "column2" ], 1 )
				.writeToCsv( workbook=wb, filepath=tempCsvPath, overwrite=true, includeHeaderRow=false );
			expect( FileRead( tempCsvPath ) ).toBe( expectedCsv );
			// move header row down one
			s.shiftRows( wb, 1, 3, 1 )
				.writeToCsv( workbook=wb, filepath=tempCsvPath, overwrite=true, includeHeaderRow=false, headerRow=2 );
			expect( FileRead( tempCsvPath ) ).toBe( expectedCsv );
		})
	})

	describe( "writeToCsv throws an exception if", ()=>{

		it( "the path exists and overwrite is false", ()=>{
			FileWrite( tempCsvPath, "" );
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.writeToCsv( wb, tempCsvPath, false );
				}).toThrow( type="cfsimplicity.spreadsheet.fileAlreadyExists" );
			})
		})

	})

})	
</cfscript>