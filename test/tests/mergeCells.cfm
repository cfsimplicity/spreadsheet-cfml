<cfscript>
describe( "mergeCells tests",function(){

	describe( "mergeCells exceptions",function(){

		beforeEach( function(){
			variables.workbook = s.new();
		});

		it( "Throws an exception if startRow OR startColumn is less than 1",function() {
			expect( function(){
				s.mergeCells( workbook,0,0,1,2 );
			}).toThrow( regex="Invalid" );
			expect( function(){
				s.mergeCells( workbook,1,2,0,0 );
			}).toThrow( regex="Invalid" );
		});

		it( "Throws an exception if endRow/endColumn is less than startRow/startColumn",function() {
			expect( function(){
				s.mergeCells( workbook,2,1,1,2 );
			}).toThrow( regex="Invalid" );
			expect( function(){
				s.mergeCells( workbook,1,2,2,1 );
			}).toThrow( regex="Invalid" );
		});

	});	
	
});	
</cfscript>