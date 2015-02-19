<cfscript>
describe( "formatColumns tests",function(){

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