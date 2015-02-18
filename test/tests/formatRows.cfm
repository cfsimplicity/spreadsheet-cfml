<cfscript>
describe( "formatRows tests",function(){

	describe( "formatRows exceptions",function(){

		it( "Throws an exception if the range is invalid",function() {
			expect( function(){
				workbook = s.new();
				format = { font="Courier" };
				s.formatRows( workbook,format,"a-b" );
			}).toThrow( regex="Invalid range" );
		});

	});

});	
</cfscript>