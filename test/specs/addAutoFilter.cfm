<cfscript>
describe( "addAutoFilter", function(){

	beforeEach( function(){
		var data = QueryNew( "Header1,Header2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
		var xls = s.workbookFromQuery( data );
		var xlsx = s.workbookFromQuery( data=data, xmlformat=true );
		variables.workbooks = [ xls, xlsx ];
	});

	it( "Doesn't error when passing valid arguments", function() {
		workbooks.Each( function( wb ){
			s.addAutoFilter( wb, "A1:B1" );
			// default to all cols in first row if no row range passed
			s.addAutoFilter( wb );
			// allow row to be specified instead of range
			s.addAutoFilter( workbook=wb, row=2 );
		});
		
	});

	it( "Doesn't error when passing valid arguments with extra trailing/leading space", function() {
		workbooks.Each( function( wb ){
			s.addAutoFilter( wb, " A1:B1 " );
		});
	});

});	
</cfscript>