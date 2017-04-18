<cfscript>
describe( "formatColumn",function(){

	it( "can format a column containing more than 4009 rows",function(){
		var path=ExpandPath( "/root/test/files/4010-rows.xls" );
		var workbook=s.read( src=path );
		var format={ italic="true" };
		s.formatColumn( workbook,format,1 );
	});

	describe( "formatColumn throws an exception if",function(){

		it( "the column is 0 or below",function(){
			expect( function(){
				workbook = s.new();
				format = { italic="true" };
				s.formatColumn( workbook,format,0 );
			}).toThrow( regex="Invalid column" );
		});

	});

});	
</cfscript>