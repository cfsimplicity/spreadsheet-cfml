<cfscript>
describe( "write tests",function(){

	it( "Writes a spreadsheet object correctly",function() {
		data = QueryNew( "column1,column2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
		workbook = s.new();
		s.addRows( workbook,data );
		s.write( workbook,tempXlsPath,true );
		expected = data;
		actual = s.read( src=tempXlsPath,format="query" );
		expect( actual ).toBe( expected );
	});

	describe( "write exceptions",function(){

		it( "Throws an exception if the path exists and overwrite is false",function() {
			FileWrite( tempXlsPath,"test" );
			workbook = s.new();
			expect( function(){
				s.write( workbook,tempXlsPath,false );
			}).toThrow( message="File already exists" );
		});

	});	
	
});	
</cfscript>