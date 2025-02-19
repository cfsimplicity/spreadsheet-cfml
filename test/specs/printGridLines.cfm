<cfscript>
describe( "printGridLines", ()=>{

	beforeEach( ()=>{
		var columnData = [ "a", "b", "c" ];
		var rowData = [ columnData, columnData, columnData ];
		var data = QueryNew( "c1,c2,c3", "VarChar,VarChar,VarChar", rowData );
		variables.xls = s.workbookFromQuery( data, false );
		variables.xlsx = s.workbookFromQuery( data=data, addHeaderRow=false, xmlformat=true );
		variables.workbooks = [ xls, xlsx ];
	})

	it( "can be added", ()=>{
		workbooks.Each( ( wb )=>{
			s.addPrintGridLines( wb );
			expect( s.getSheetHelper().getActiveSheet( wb ).isPrintGridlines() ).toBeTrue();
		})
	})

	it( "can be removed", ()=>{
		workbooks.Each( ( wb )=>{
			s.addPrintGridLines( wb )
				.removePrintGridLines( wb );
			expect( s.getSheetHelper().getActiveSheet( wb ).isPrintGridlines() ).toBeFalse();
		})
	})

	it( "addPrintGridLines and removePrintGridLines are chainable", ()=>{
		workbooks.Each( ( wb )=>{
			s.newChainable( wb ).addPrintGridLines();
			expect( s.getSheetHelper().getActiveSheet( wb ).isPrintGridlines() ).toBeTrue();
			s.newChainable( wb ).removePrintGridLines();
			expect( s.getSheetHelper().getActiveSheet( wb ).isPrintGridlines() ).toBeFalse();
		})
	})

})	
</cfscript>