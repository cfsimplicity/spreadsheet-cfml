<cfscript>
describe( "formatColumns tests",function(){

	it( "can format columns in a spreadsheet containing more than 4000 rows",function(){
		var path=ExpandPath( "/root/test/files/4001.xls" );
		var workbook=s.read( src=path );
		var format={ italic="true" };
		s.formatColumns( workbook,format,"1-2" );
	});

	describe( "formatColumns exceptions",function(){

		it( "Throws an exception if the range is invalid",function() {
			expect( function(){
				workbook = s.new();
				format = { font="Courier" };
				s.formatColumns( workbook,format,"a-b" );
			}).toThrow( regex="Invalid range" );
		});

	});

});	
</cfscript>