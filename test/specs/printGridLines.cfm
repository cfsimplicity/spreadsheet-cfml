<cfscript>
describe( "printGridLines", function(){

	beforeEach( function(){
		var columnData = [ "a", "b", "c" ];
		var rowData = [ columnData, columnData, columnData ];
		var data = QueryNew( "c1,c2,c3", "VarChar,VarChar,VarChar", rowData );
		variables.xls = s.workbookFromQuery( data, false );
		variables.xlsx = s.workbookFromQuery( data=data, addHeaderRow=false, xmlformat=true );
		variables.workbooks = [ xls, xlsx ];
		makePublic( s, "getActiveSheet" );
	});

	it( "can be added", function(){
		workbooks.Each( function( wb ){
			s.addPrintGridLines( wb );
			expect( s.getActiveSheet( wb ).isPrintGridlines() ).toBeTrue();
		});
	});

	it( "can be removed", function(){
		workbooks.Each( function( wb ){
			s.addPrintGridLines( wb );
			s.removePrintGridLines( wb );
			expect( s.getActiveSheet( wb ).isPrintGridlines() ).toBeFalse();
		});
	});

});	
</cfscript>