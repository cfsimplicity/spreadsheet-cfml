<cfscript>
describe( "formatColumn tests",function(){

	describe( "formatColumn exceptions",function(){

		it( "Throws an exception if the column is 0 or below",function() {
			expect( function(){
				workbook = s.new();
				format = { italic="true" };
				s.formatColumn( workbook,format,0 );
			}).toThrow( regex="Invalid column" );
		});

	});

});	
</cfscript>