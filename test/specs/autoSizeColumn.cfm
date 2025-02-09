<cfscript>
describe( "autoSizeColumn", ()=>{

	beforeEach( ()=>{
		var data = QueryNew( "First,Last", "VarChar,VarChar", [ [ "a", "abracadabraabracadabra" ] ] );
		var xls = s.workbookFromQuery( data );
		var xlsx = s.workbookFromQuery( data=data, xmlFormat=true );
		variables.workbooks = [ xls, xlsx ];
	})

	it( "Doesn't error when passing valid arguments", ()=>{
		workbooks.Each( ( wb )=>{
			s.autoSizeColumn( wb, 2 );
		})
	})

	it( "Is chainable", ()=>{
		workbooks.Each( ( wb )=>{
			s.newChainable( wb ).autoSizeColumn( 2 );
		})
	})

	it( "Doesn't error if the workbook is SXSSF", ()=>{
		var data = QueryNew( "First,Last", "VarChar,VarChar", [ [ "a", "abracadabraabracadabra" ] ] );
		var workbook = s.newStreamingXlsx();
		s.addRows( local.workbook, data )
			.autoSizeColumn( local.workbook, 2 );
	})

	it( "Throws a helpful exception if column argument is invalid", ()=>{
		workbooks.Each( ( wb )=>{
			expect( ()=>{
				s.autoSizeColumn( wb, -1 );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidColumnArgument" );
		})
	})

})	
</cfscript>