<cfscript>
describe( "read: format=csv", ()=>{

	it( "Can return a CSV string from an Excel file", ()=>{
		var path = getTestFilePath( "test.xls" );
		var expected = 'a,b#newline#1,2015-04-01 00:00:00#newline#2015-04-01 01:01:01,2#newline#';
		var actual = s.read( src=path, format="csv" );
		expect( actual ).toBe( expected );
		expected = 'a,b#newline#a,b#newline#1,2015-04-01 00:00:00#newline#2015-04-01 01:01:01,2#newline#';
		actual = s.read( src=path, format="csv", headerRow=1, includeHeaderRow=true );
		expect( actual ).toBe( expected );
	})

	it( "Escapes double-quotes in string values when reading to CSV", ()=>{
		var data = QueryNew( "column1", "VarChar", [ [ 'a "so-called" test' ] ] );
		var workbook = s.new();
		s.addRows( workbook, data )
			.write( workbook, tempXlsPath, true );
		var expected = '"a ""so-called"" test"#newline#';
		var actual = s.read( src=tempXlsPath, format="csv" );
		expect( actual ).toBe( expected );
	})

	it( "Accepts a custom delimiter when generating CSV", ()=>{
		var path = getTestFilePath( "test.xls" );
		var expected = 'a|b#newline#1|2015-04-01 00:00:00#newline#2015-04-01 01:01:01|2#newline#';
		var actual = s.read( src=path, format="csv", csvDelimiter="|" );
		expect( actual ).toBe( expected );
	})

})
</cfscript>
