<cfscript>
describe( "shiftRows", ()=>{

	beforeEach( ()=>{
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		variables.rowData = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
	})

	it( "Shifts rows down if offset is positive", ()=>{
		workbooks.Each( ( wb )=>{
			s.addRows( wb, rowData )
				.shiftRows( wb, 1, 1, 1 );
			var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "" ], [ "a", "b" ] ] );
			var actual = s.getSheetHelper().sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		})
	})

	it( "Shifts rows up if offset is negative", ()=>{
		workbooks.Each( ( wb )=>{
			s.addRows( wb, rowData )
				.shiftRows( wb, 2, 2, -1 );
			var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "c", "d" ] ] );
			var actual = s.getSheetHelper().sheetToQuery( wb );
			expect( actual ).toBe( expected );
		})
	})

	it( "is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			s.newChainable( wb )
				.addRows( rowData )
				.shiftRows( 1, 1, 1 );
			var expected = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "", "" ], [ "a", "b" ] ] );
			var actual = s.getSheetHelper().sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		})
	})

})	
</cfscript>