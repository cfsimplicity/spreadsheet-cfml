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
	});

	it( "Doesn't error when passing valid arguments with extra trailing/leading space",function() {
		s.addAutoFilter( xls, " A1:B1 " );
		s.addAutoFilter( xlsx, " A1:B1 " );
	});

	it( "Throws a helpful exception if range argument is present but empty",function() {
		expect( function(){
			s.addAutoFilter( xls, "" );
		}).toThrow( regex="Empty cellRange argument" );
	});

});	
</cfscript>