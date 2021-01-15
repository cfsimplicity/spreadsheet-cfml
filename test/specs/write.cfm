<cfscript>
describe( "write", function(){

	beforeEach( function(){
		sleep( 5 );// allow time for file operations to complete
	});

	it( "Writes an XLS object correctly", function(){
		data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a","b" ], [ "c","d" ] ] );
		var workbook = s.newXls();
		s.addRows( workbook, data );
		s.write( workbook, tempXlsPath, true );
		var expected = data;
		var actual = s.read( src=tempXlsPath, format="query" );
		expect( actual ).toBe( expected );
	});

	it( "Writes an XLSX object correctly", function(){
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
		var workbook = s.newXlsx();
		s.addRows( workbook, data );
		s.write( workbook, tempXlsxPath, true );
		var expected = data;
		var actual = s.read( src=tempXlsxPath, format="query" );
		expect( actual ).toBe( expected );
	});

	it( "Writes a streaming XLSX object without error", function(){
		var rows = [];
		for( i=1; i <= 500; i++ ){
			rows.append( { column1=i, column2="test" } );
		}
		var data = QueryNew( "column1,column2", "Integer,Varchar", rows );
		var workbook = s.newStreamingXlsx();
		s.addRows( workbook, data );
		s.write( workbook, tempXlsxPath, true );
		var expected = data;
		var actual = s.read( src=tempXlsxPath, format="query" );
		expect( actual ).toBe( expected );
	});

	it( "Writes a streaming XLSX object with a custom window size without error", function(){
		var rows = [];
		for( i=1; i <= 500; i++ ){
			rows.append( { column1=i, column2="test" } );
		}
		var data = QueryNew( "column1,column2", "Integer,Varchar", rows );
		var workbook = s.newStreamingXlsx( streamingWindowSize=2 );
		s.addRows( workbook, data );
		s.write( workbook, tempXlsxPath, true );
		var expected = data;
		var actual = s.read( src=tempXlsxPath, format="query" );
		expect( actual ).toBe( expected );
	});
		
	it( "Can write an XLSX file encrypted with a password", function(){
		var data = QueryNew( "column1", "VarChar", [ [ "secret" ] ] );
		var workbook = s.newXlsx();
		s.addRows( workbook,data );
		s.write( workbook=workbook, filepath=tempXlsxPath, overwrite=true, password="pass" );
		var expected = data;
		var actual = s.read( src=tempXlsxPath, format="query", password="pass" );
		expect( actual ).toBe( expected );
	});

	describe( "write throws an exception if", function(){

		it( "the path exists and overwrite is false", function(){
			FileWrite( tempXlsPath, "test" );
			var workbook = s.new();
			expect( function(){
				s.write( workbook, tempXlsPath, false );
			}).toThrow( message="File already exists" );
		});

	});	

	afterEach( function(){
		if( FileExists( variables.tempXlsPath ) ) FileDelete( variables.tempXlsPath );
		if( FileExists( variables.tempXlsxPath ) ) FileDelete( variables.tempXlsxPath );
	});
	
});	
</cfscript>