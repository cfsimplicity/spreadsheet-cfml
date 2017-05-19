<cfscript>
describe( "writeFileFromQuery",function(){

	beforeEach( function(){
		sleep( 5 );// allow time for file operations to complete
		variables.query = QueryNew( "Header1,Header2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
	});

	it( "Writes a file from a query",function() {
		s.writeFileFromQuery( query,tempXlsPath,true );
		expected = query;
		actual = s.read( src=tempXlsPath,format="query",headerRow=1 );
		expect( actual ).toBe( expected );
	});

	it( "Writes an OOXML file if extension is .xlsx",function() {
		path = tempXlsxPath;
		s.writeFileFromQuery( query,path,true );
		workbook	=	s.read( path );
		expect( workbook.getClass().name ).toBe( "org.apache.poi.xssf.usermodel.XSSFWorkbook" );
	});

	it( "Writes an OOXML file if extension is .xls but xmlFormat is true",function() {
		convertedPath = tempXlsPath & "x";
		s.writeFileFromQuery( data=query,filepath=tempXlsPath,overwrite=true,xmlFormat=true );
		workbook	=	s.read( convertedPath );
		expect( workbook.getClass().name ).toBe( "org.apache.poi.xssf.usermodel.XSSFWorkbook" );
		if( FileExists( convertedPath ) )
			FileDelete( convertedPath );
	});

});	
</cfscript>