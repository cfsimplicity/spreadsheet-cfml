<cfscript>
describe( "writeFileFromQuery tests",function(){

	beforeEach( function(){
		variables.data = QueryNew( "Header1,Header2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
	});

	it( "Writes a file from a query",function() {
		s.writeFileFromQuery( data,tempXlsPath,true );
		expected = data;
		actual = s.read( src=tempXlsPath,format="query",headerRow=1 );
		expect( actual ).toBe( expected );
	});

	it( "Writes an OOXML file if extension is .xlsx",function() {
		path = ExpandPath( "/root/test/test.xlsx" );
		s.writeFileFromQuery( data,path,true );
		workbook	=	s.read( path );
		expect( workbook.getClass().name ).toBe( "org.apache.poi.xssf.usermodel.XSSFWorkbook" );
		FileDelete( path );
	});

	it( "Writes an OOXML file if extension is .xls but xmlFormat is true",function() {
		convertedPath = tempXlsPath & "x";
		s.writeFileFromQuery( data=data,filepath=tempXlsPath,overwrite=true,xmlFormat=true );
		workbook	=	s.read( convertedPath );
		expect( workbook.getClass().name ).toBe( "org.apache.poi.xssf.usermodel.XSSFWorkbook" );
		if( FileExists( convertedPath ) )
			FileDelete( convertedPath );
	});

});	
</cfscript>