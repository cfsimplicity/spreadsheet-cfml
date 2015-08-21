<cfscript>
describe( "rich text formatting tests",function(){

	beforeEach( function(){
		path = ExpandPath( "/root/test/files/formatting.xls" );
		actual = s.read( src=path,format="query",includeRichTextFormatting="true" );
	});

	it( "parses line 1: whole cell unformatted",function() {
		expected = 'unformatted';
		expect( actual.column1[ 1 ] ).toBe( expected );
	});

	it( "parses line 2: whole cell bold",function() {
		expected = '<span style="font-weight:bold;">bold</span>';
		expect( actual.column1[ 2 ] ).toBe( expected );
	});

	it( "parses line 3: whole cell red",function() {
		expected = '<span style="color:##ff3333;">red</span>';
		expect( actual.column1[ 3 ] ).toBe( expected );
	});

	it( "parses line 4: whole cell italic",function() {
		expected = '<span style="font-style:italic;">italic</span>';
		expect( actual.column1[ 4 ] ).toBe( expected );
	});

	it( "parses line 5: whole cell strike",function() {
		expected = '<span style="text-decoration:line-through;">strike</span>';
		expect( actual.column1[ 5 ] ).toBe( expected );
	});

	it( "parses line 6: whole cell underline",function() {
		expected = '<span style="text-decoration:underline;">underline</span>';
		expect( actual.column1[ 6 ] ).toBe( expected );
	});

	it( "parses line 7: whole cell bold red italic strike underline",function() {
		expected = '<span style="font-weight:bold;color:##ff3333;font-style:italic;text-decoration:line-through underline;">bold red italic strike underline</span>';
		expect( actual.column1[ 7 ] ).toBe( expected );
	});

	it( "parses line 8: unformatted + bold",function() {
		expected = 'unformatted<span style="font-weight:bold;">bold</span>';
		expect( actual.column1[ 8 ] ).toBe( expected );
	});

	it( "parses line 9: unformatted + red",function() {
		expected = 'unformatted<span style="color:##ff3333;">red</span>';
		expect( actual.column1[ 9 ] ).toBe( expected );
	});

	it( "parses line 10: unformatted + italic",function() {
		expected = 'unformatted<span style="font-style:italic;">italic</span>';
		expect( actual.column1[ 10 ] ).toBe( expected );
	});

	it( "parses line 11: unformatted + strike",function() {
		expected = 'unformatted<span style="text-decoration:line-through;">strike</span>';
		expect( actual.column1[ 11 ] ).toBe( expected );
	});

	it( "parses line 12: unformatted underline",function() {
		expected = 'unformatted<span style="text-decoration:underline;">underline</span>';
		expect( actual.column1[ 12 ] ).toBe( expected );
	});

	it( "parses line 13: unformatted + bold red italic strike underline",function() {
		expected = 'unformatted<span style="font-weight:bold;color:##ff3333;font-style:italic;text-decoration:line-through underline;">bold red italic strike underline</span>';
		expect( actual.column1[ 13 ] ).toBe( expected );
	});

	it( "parses line 14: unformatted + shadow (= unsupported style)",function() {
		expected = 'unformattedShadow';
		expect( actual.column1[ 14 ] ).toBe( expected );
	});

	it( "parses line 15: bold + unformatted",function() {
		expected = '<span style="font-weight:bold;">bold</span><span style="font-weight:normal;">unformatted</span>';
		expect( actual.column1[ 15 ] ).toBe( expected );
	});

	it( "parses line 16: bold + red + italic + strike + underline",function() {
		expected = '<span style="font-weight:bold;">bold</span><span style="font-weight:normal;color:##ff3333;">red</span><span style="font-weight:normal;font-style:italic;">italic</span><span style="font-weight:normal;text-decoration:line-through;">strike</span><span style="font-weight:normal;text-decoration:underline;">underline</span>';
		expect( actual.column1[ 16 ] ).toBe( expected );
	});

});
</cfscript>