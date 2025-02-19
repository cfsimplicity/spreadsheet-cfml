<cfscript>
describe( "writeFileFromQuery", ()=>{

	beforeEach( ()=>{
		variables.query = QueryNew( "Header1,Header2","VarChar,VarChar",[ [ "a","b" ],[ "c","d" ] ] );
	})

	it( "Writes a file from a query", ()=>{
		s.writeFileFromQuery( query, tempXlsPath, true );
		var expected = query;
		var actual = s.read( src=tempXlsPath, format="query", headerRow=1 );
		expect( actual ).toBe( expected );
	})

	it( "Writes an XLSX file if extension is .xlsx", ()=>{
		var path = tempXlsxPath;
		s.writeFileFromQuery( query, path, true );
		var workbook	=	s.read( path );
		expect( workbook.getClass().name ).toBe( "org.apache.poi.xssf.usermodel.XSSFWorkbook" );
	})

	it( "Writes an XLSX file if extension is .xls but xmlFormat is true", ()=>{
		var convertedPath = tempXlsPath & "x";
		s.writeFileFromQuery( data=query, filepath=tempXlsPath, overwrite=true, xmlFormat=true );
		var workbook	=	s.read( convertedPath );
		expect( workbook.getClass().name ).toBe( "org.apache.poi.xssf.usermodel.XSSFWorkbook" );
		if( FileExists( convertedPath ) )
			FileDelete( convertedPath );
	})

})	
</cfscript>