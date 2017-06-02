<cfscript>
describe( "printGridLines",function(){

	beforeEach( function(){
		var columnData = [ "a", "b", "c" ];
		var rowData = [ columnData, columnData, columnData ];
		variables.data = QueryNew( "c1,c2,c3", "VarChar,VarChar,VarChar", rowData );
		variables.xls = s.workbookFromQuery( data, false );
		variables.xlsx = s.workbookFromQuery( data=data, addHeaderRow=false, xmlformat=true );
		makePublic( s, "getActiveSheet" );
		variables.xlsSheet = s.getActiveSheet( xls );
		variables.xlsxSheet = s.getActiveSheet( xlsx );
	});

	it( "can be added",function() {
		s.addPrintGridLines( xls );
		s.addPrintGridLines( xlsx );
		expect( xlsSheet.isPrintGridlines() ).toBeTrue();
		expect( xlsxSheet.isPrintGridlines() ).toBeTrue();
	});

	it( "can be removed",function() {
		s.addPrintGridLines( xls );
		s.addPrintGridLines( xlsx );
		s.removePrintGridLines( xls );
		s.removePrintGridLines( xlsx );
		expect( xlsSheet.isPrintGridlines() ).toBeFalse();
		expect( xlsxSheet.isPrintGridlines() ).toBeFalse();
	});

});	
</cfscript>