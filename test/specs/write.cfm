<cfscript>
describe( "write", function(){

	beforeEach( function(){
		sleep( 5 );// allow time for file operations to complete
	});

	it( "Writes an XLS object correctly", function() {
		data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a","b" ], [ "c","d" ] ] );
		workbook = s.newXls();
		s.addRows( workbook, data );
		s.write( workbook, tempXlsPath, true );
		expected = data;
		actual = s.read( src=tempXlsPath, format="query" );
		expect( actual ).toBe( expected );
	});

	it( "Writes an XLSX object correctly", function() {
		data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a","b" ], [ "c","d" ] ] );
		workbook = s.newXlsx();
		s.addRows( workbook, data );
		s.write( workbook, tempXlsxPath, true );
		expected = data;
		actual = s.read( src=tempXlsxPath, format="query" );
		expect( actual ).toBe( expected );
	});

	it( "Writes a streaming XLSX object without error", function(){
		var rows = [];
		for( i=1; i <= 500; i++ ){
			rows.append( { column1=i, column2="test" } );
		}
		var data = QueryNew( "column1,column2", "Integer,Varchar", rows );
		workbook = s.newStreamingXlsx();
		s.addRows( workbook, data );
		s.write( workbook, tempXlsxPath, true );
		expected = data;
		actual = s.read( src=tempXlsxPath, format="query" );
		expect( actual ).toBe( expected );
	});

	it( "Writes a streaming XLSX object with a custom window size without error", function(){
		var rows = [];
		for( i=1; i <= 500; i++ ){
			rows.append( { column1=i, column2="test" } );
		}
		var data = QueryNew( "column1,column2", "Integer,Varchar", rows );
		workbook = s.newStreamingXlsx( streamingWindowSize=2 );
		s.addRows( workbook, data );
		s.write( workbook, tempXlsxPath, true );
		expected = data;
		actual = s.read( src=tempXlsxPath, format="query" );
		expect( actual ).toBe( expected );
	});

	describe( "write throws an exception if", function(){

		it( "the path exists and overwrite is false", function() {
			FileWrite( tempXlsPath, "test" );
			workbook = s.new();
			expect( function(){
				s.write( workbook, tempXlsPath, false );
			}).toThrow( message="File already exists" );
		});

	});	
	
});	
</cfscript>