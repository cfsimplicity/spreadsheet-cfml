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
		expect( actual ).toBe( expected );
	});
	it( "parses the complex file",function() {
		path = ExpandPath( "/root/test/files/format.xls" );
		expected = querySim(
			"column1
			£99 <span style=""text-decoration: line-through;"">£55</span>
			£99 <span style=""text-decoration: line-through;"">£55</span>
			£99 <span style=""text-decoration: line-through;"">£55</span>
			£99 <span style=""text-decoration: line-through;"">£55</span>
			£99 <span style=""text-decoration: line-through;"">£55</span>
			£99 <span style=""text-decoration: line-through;"">£55</span>
			£99 <span style=""text-decoration: line-through;"">£55</span>
			£99 <span style=""text-decoration: line-through;"">£55</span>
			£99 <span style=""text-decoration: line-through;"">£55</span>
			£99 <span style=""text-decoration: line-through;"">£55</span>
			£99 <span style=""text-decoration: line-through;"">£55</span>");
		actual = s.read( src=path,format="query",exportRichText="true",includeHeaderRow=true );
		expect( actual ).toBe( expected );
	});

});1
</cfscript>