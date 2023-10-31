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

	it( "outputs date objects using the instance's specified DATETIME format", function(){
		var nowAsText = DateTimeFormat( Now(), s.getDateFormats().DATETIME );
		var data = QueryNew( "column1", "Timestamp", [ [ ParseDateTime( nowAsText ) ] ] );
		var expected = '#nowAsText#';
		expect( s.queryToCsv( data ) ).toBe( expected );
	});

	it( "does NOT treat date strings as date objects to be formatted using the DATETIME format", function(){
		var dateString = "2022-12-18";
		var data = QueryNew( "column1", "VarChar", [ [ dateString ] ] );
		var expected = '#dateString#';
		expect( s.queryToCsv( data ) ).toBe( expected );
	});

	it(
		title="can process rows in parallel if the engine supports it"
		,body=function(){
			//can't test if using threads, just that there are no errors
			var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "a" ], [ "a", "a" ] ] );
			var expected = 'a,a#crlf#a,a';//same values because order is not guaranteed
			expect( s.queryToCsv( query=data, threads=2 ) ).toBe( expected );
		}
		,skip=function(){
			//20231031: ACF 2021 and 2023 won't run the whole suite if this test is included: testbox errors thrown
			//running just the queryToCsv tests works fine though. Lucee is fine with the whole suite.
			return s.getIsACF();
		}
	);

	describe( "queryToCsv throws an exception if", function(){

		it(
			title="parallel threads are specified and the engine does not support it"
			,body=function(){
				expect( function(){
					s.queryToCsv( query=data, threads=2 );
				}).toThrow( type="cfsimplicity.spreadsheet.parallelOptionNotSupported" );
			}
			,skip=function(){
				return s.engineSupportsParallelLoopProcessing();
			}
		);

	});

});
</cfscript>