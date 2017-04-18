<cfscript>
describe( "formatRows",function(){

	describe( "formatRows throws an exception if",function(){

		it( "the range is invalid",function() {
			expect( function(){
				workbook = s.new();
				format = { font="Courier" };
				s.formatRows( workbook,format,"a-b" );
			}).toThrow( regex="Invalid range" );
		});

	});

});	
</cfscript>