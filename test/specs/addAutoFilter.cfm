<cfscript>
describe( "addAutoFilter",function(){

	beforeEach( function(){
		var data = QueryNew( "Header1,Header2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
		variables.xls = s.workbookFromQuery( data );
		variables.xlsx = s.workbookFromQuery( data=data, xmlformat=true );
	});

	it( "Doesn't error when passing valid arguments",function() {
		s.addAutoFilter( xls, "A1:B1" );
		s.addAutoFilter( xlsx, "A1:B1" );
		// default to all cols in first row if no row range passed
		s.addAutoFilter( xls );
		s.addAutoFilter( xlsx );
		// allow row to be specified instead of range
		s.addAutoFilter( workbook = xls, row = 2 );
		s.addAutoFilter( workbook = xlsx, row = 2 );
	});

	it( "Doesn't error when passing valid arguments with extra trailing/leading space",function() {
		s.addAutoFilter( xls, " A1:B1 " );
		s.addAutoFilter( xlsx, " A1:B1 " );
	});

});	
</cfscript>