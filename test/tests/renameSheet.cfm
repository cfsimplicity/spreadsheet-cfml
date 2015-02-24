<cfscript>
describe( "renameSheet tests",function(){

	it( "Renames the specified sheet",function() {
		var workbook = s.new();
		s.renameSheet( workbook,"test",1 );
		makePublic( s,"sheetExists" )
		expect( s.sheetExists( workbook,"test" ) ).toBeTrue();
	});


	describe( "renameSheet exceptions",function(){

		it( "Throws an exception if the new sheet name already exists",function() {
			expect( function(){
				var workbook = s.new();
				s.createSheet( workbook,"test" );
				s.createSheet( workbook,"test2" );
				s.renameSheet( workbook,"test2",2 );
			}).toThrow( regex="Invalid Sheet Name" );
		});

	});	

});	
</cfscript>