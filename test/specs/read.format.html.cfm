<cfscript>
describe( "read: format=html", ()=>{

	it( "Can return HTML table rows from an Excel file", ()=>{
		var path = getTestFilePath( "test.xls" );
		var actual = s.read( src=path, format="html" );
		var expected = "<tbody><tr><td>a</td><td>b</td></tr><tr><td>1</td><td>2015-04-01 00:00:00</td></tr><tr><td>2015-04-01 01:01:01</td><td>2</td></tr></tbody>";
		expect( actual ).toBe( expected );
		actual = s.read( src=path, format="html", headerRow=1 );
		expected = "<tbody><tr><td>1</td><td>2015-04-01 00:00:00</td></tr><tr><td>2015-04-01 01:01:01</td><td>2</td></tr></tbody>";
		expect( actual ).toBe( expected );
		actual = s.read( src=path, format="html", headerRow=1, includeHeaderRow=true );
		expected="<thead><tr><th>a</th><th>b</th></tr></thead><tbody><tr><td>a</td><td>b</td></tr><tr><td>1</td><td>2015-04-01 00:00:00</td></tr><tr><td>2015-04-01 01:01:01</td><td>2</td></tr></tbody>";
		expect( actual ).toBe( expected );
	})
	
})
</cfscript>
