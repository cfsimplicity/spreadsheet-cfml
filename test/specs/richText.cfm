<cfscript>
describe( "rich text format tests",function(){

	it( "Can read the simple XLS file",function() {
		path = ExpandPath( "/root/test/files/format-simple.xls" );
		workbook = s.read( src=path,format="query",exportRichText="true");
	});

	it( "parses the simple file",function() {
		path = ExpandPath( "/root/test/files/format-simple.xls" );
		expected = querySim(
			"column1
			£99 <span style=""text-decoration: line-through;"">£55</span>");
		actual = s.read( src=path,format="query",exportRichText="true",includeHeaderRow=true );
		expect( actual.column1 ).toBe( expected.column1 );
	});
	it( "parses the complex file line 1",function() {
		path = ExpandPath( "/root/test/files/format.xls" );
		expected = "£99 <span style=""text-decoration: line-through;"">£55</span>";
		actual = s.read( src=path,format="query",exportRichText="true",includeHeaderRow=true,rows=1 );
		expect( actual.column1 ).toBe( expected );
	});
	it( "parses the complex file line 2",function() {
		path = ExpandPath( "/root/test/files/format.xls" );
		expected = "£99 <span style=""font-weight: bold;"">£56</span>";
		actual = s.read( src=path,format="query",exportRichText="true",includeHeaderRow=true,rows=2 );
		expect( actual.column1 ).toBe( expected );
	});
	it( "parses the complex file line 3",function() {
		path = ExpandPath( "/root/test/files/format.xls" );
		expected = "<span style=""text-decoration: line-through;"">£99</span><span style=""text-decoration: none;""> £55</span>";
		actual = s.read( src=path,format="query",exportRichText="true",includeHeaderRow=true,rows=3 );
		expect( actual.column1 ).toBe( expected );
	});
	it( "parses the complex file line 4",function() {
		path = ExpandPath( "/root/test/files/format.xls" );
		expected = "<span style=""font-style: italic;"">£99</span><span style=""font-weight:bold;""> £57</span>";
		actual = s.read( src=path,format="query",exportRichText="true",includeHeaderRow=true,rows=4 );
		expect( actual.column1 ).toBe( expected );
	});
	it( "parses the complex file line 5",function() {
		path = ExpandPath( "/root/test/files/format.xls" );
		expected = "<span style=""font-style: italic;"">£99<span style=""text-decoration: line-through;""> £58</span>";
		actual = s.read( src=path,format="query",exportRichText="true",includeHeaderRow=true,rows=5 );
		expect( actual.column1 ).toBe( expected );
	});
	it( "parses the complex file line 6",function() {
		path = ExpandPath( "/root/test/files/format.xls" );
		expected = "<span style=""font-weight:bold;"">£99 £59</span>";
		actual = s.read( src=path,format="query",exportRichText="true",includeHeaderRow=true,rows=6 );
		expect( actual.column1 ).toBe( expected );
	});
	it( "parses the complex file line 7 (unchanged)",function() {
		path = ExpandPath( "/root/test/files/format.xls" );
		expected = "unchanged because unformatted ";
		actual = s.read( src=path,format="query",exportRichText="true",includeHeaderRow=true,rows=7 );
		expect( actual.column1 ).toBe( expected );
	});

});1
</cfscript>