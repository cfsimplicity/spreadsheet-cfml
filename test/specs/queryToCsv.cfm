<cfscript>
describe( "queryToCsv", ()=>{

	beforeEach( ()=>{
		variables.data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
	})

	it( "converts a basic query to csv without a header row by default", ()=>{
		var expected = 'a,b#newline#c,d#newline#';
		expect( s.queryToCsv( data ) ).toBe( expected );
	})

	it( "uses the query columns as the header row if specified", ()=>{
		var expected = 'column1,column2#newline#a,b#newline#c,d#newline#';
		expect( s.queryToCsv( data, true ) ).toBe( expected );
	})

	it( "allows an alternative to the default comma delimiter", ()=>{
		var expected = 'a|b#newline#c|d#newline#';
		expect( s.queryToCsv( query=data, delimiter="|" ) ).toBe( expected );
	})

	it( "allows tabs to be specified as the delimiter in a number of ways", ()=>{
		var expected = 'a#Chr( 9 )#b#newline#c#Chr( 9 )#d#newline#';
		var validTabValues = [ "#Chr( 9 )#", "\t", "tab", "TAB" ];
		for( var value in validTabValues ){
			expect( s.queryToCsv( query=data, delimiter=value ) ).toBe( expected );
		}
	})

	it( "can handle an embedded delimiter", ()=>{
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a,a", "b" ], [ "c", "d" ] ] );
		var expected = '"a,a",b#newline#c,d#newline#';
		expect( s.queryToCsv( data ) ).toBe( expected );
	})

	it( "can handle an embedded double-quote", ()=>{
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a""a", "b" ], [ "c", "d" ] ] );
		var expected = '"a""a",b#newline#c,d#newline#';
		expect( s.queryToCsv( data ) ).toBe( expected );
	})

	it( "can handle an embedded carriage return", ()=>{
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a#newline#a", "b" ], [ "c", "d" ] ] );
		var expected = '"a#newline#a",b#newline#c,d#newline#';
		expect( s.queryToCsv( data ) ).toBe( expected );
	})

	it( "outputs date objects using the instance's specified DATETIME format", ()=>{
		var nowAsText = DateTimeFormat( Now(), s.getDateFormats().DATETIME );
		var data = QueryNew( "column1", "Timestamp", [ [ ParseDateTime( nowAsText ) ] ] );
		var expected = '#nowAsText##newline#';
		expect( s.queryToCsv( data ) ).toBe( expected );
	})

	it( "does NOT treat date strings as date objects to be formatted using the DATETIME format", ()=>{
		var dateString = "2022-12-18";
		var data = QueryNew( "column1", "VarChar", [ [ dateString ] ] );
		var expected = '#dateString##newline#';
		expect( s.queryToCsv( data ) ).toBe( expected );
	})

	it(
		title="can process rows in parallel if the engine supports it"
		,body=function(){
			//can't test if using threads, just that there are no errors
			var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "a" ], [ "a", "a" ] ] );
			var expected = 'a,a#newline#a,a#newline#';//same values because order is not guaranteed
			expect( s.queryToCsv( query=data, threads=2 ) ).toBe( expected );
		}
		,skip=function(){
			//20231031: ACF 2021 and 2023 won't run the whole suite if this test is included: testbox errors thrown
			//running just the queryToCsv tests works fine though. Lucee is fine with the whole suite.
			return s.getIsACF();
		}
	);

	describe( "queryToCsv throws an exception if", ()=>{

		it(
			title="parallel threads are specified and the engine does not support it"
			,body=function(){
				expect( ()=>{
					s.queryToCsv( query=data, threads=2 );
				}).toThrow( type="cfsimplicity.spreadsheet.parallelOptionNotSupported" );
			}
			,skip=s.engineSupportsParallelLoopProcessing()
		);

	})

})
</cfscript>