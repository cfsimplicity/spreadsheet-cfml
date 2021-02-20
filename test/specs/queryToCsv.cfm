<cfscript>
describe( "queryToCsv", function(){

	beforeEach( function(){
		variables.data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
	});

	it( "converts a basic query to csv without a header row by default", function(){
		var expected = 'a,b#crlf#c,d';
		expect( s.queryToCsv( data ) ).toBe( expected );
	});

	it( "uses the query columns as the header row if specified", function(){
		var expected = 'column1,column2#crlf#a,b#crlf#c,d';
		expect( s.queryToCsv( data, true ) ).toBe( expected );
	});

	it( "allows an alternative to the default comma delimiter", function(){
		var expected = 'a|b#crlf#c|d';
		expect( s.queryToCsv( query=data, delimiter="|" ) ).toBe( expected );
	});

	it( "allows tabs to be specified as the delimiter in a number of ways", function(){
		var expected = 'a#Chr( 9 )#b#crlf#c#Chr( 9 )#d';
		var validTabValues = [ "#Chr( 9 )#", "\t", "tab", "TAB" ];
		for( var value in validTabValues ){
			expect( s.queryToCsv( query=data, delimiter=value ) ).toBe( expected );
		}
	});

	it( "can handle an embedded delimiter", function(){
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a,a", "b" ], [ "c", "d" ] ] );
		var expected = '"a,a",b#crlf#c,d';
		expect( s.queryToCsv( data ) ).toBe( expected );
	});

	it( "can handle an embedded double-quote", function(){
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a""a", "b" ], [ "c", "d" ] ] );
		var expected = '"a""a",b#crlf#c,d';
		expect( s.queryToCsv( data ) ).toBe( expected );
	});

	it( "can handle an embedded carriage return", function(){
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a#crlf#a", "b" ], [ "c", "d" ] ] );
		var expected = '"a#crlf#a",b#crlf#c,d';
		expect( s.queryToCsv( data ) ).toBe( expected );
	});

});
</cfscript>