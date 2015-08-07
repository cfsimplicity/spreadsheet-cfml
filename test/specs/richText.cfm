<cfscript>
describe( "rich text format tests",function(){

	it( "Can read a traditional XLS file",function() {
		path = ExpandPath( "/root/test/files/format.xls" );
		workbook = s.read( src=path,format="query",exportRichText="true");
	});

	it( "parses",function() {
		path = ExpandPath( "/root/test/files/format.xls" );
		expected = querySim(
			"a,b
			1|#ParseDateTime( '2015-04-01 00:00:00' )#
			#ParseDateTime( '2015-04-01 01:01:01' )#|2");
		actual = s.read( src=path,format="query",headerRow=1 );
		expect( actual ).toBe( expected );
	});

});1
</cfscript>