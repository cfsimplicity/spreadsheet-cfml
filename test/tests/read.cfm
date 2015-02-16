<cfscript>
describe( "read tests",function(){

	it( "can read a traditional XLS file",function() {
		path = ExpandPath( "/root/test/files/test.xls" );
		expected = querySim(
			"column1,column2
			a|b
			c|d");
		actual = s.read( src=path,format="query" );
		expect( actual ).toBe( expected );
	});

	it( "can read an OOXML file",function() {
		path = ExpandPath( "/root/test/files/test.xlsx" );
		expected = querySim(
			"column1,column2
			a|e
			b|f
			c|g
			I am|ooxml");
		actual = s.read( src=path,format="query" );
		expect( actual ).toBe( expected );
	});

});	
</cfscript>