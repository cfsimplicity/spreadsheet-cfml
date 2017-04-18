<cfscript>
describe( "csvToQuery",function(){

	beforeEach( function(){
		variables.basicExpectedQuery = QueryNew( "column1,column2", "", [ [ "Frumpo McNugget", "12345" ] ] );
	});

	it( "converts a basic comma delimited, double quote qualified csv string to a query", function() {
		savecontent variable="local.csv"{
			WriteOutput( '
"Frumpo McNugget",12345
		');
		};
		var actual = s.csvToQuery( csv );
		expect( actual ).toBe( basicExpectedQuery ); 
	});

	it( "can read the csv from a file", function() {
		var path = ExpandPath( "/root/test/files/test.csv" );
		//named args
		var actual = s.csvToQuery( filepath=path );
		expect( actual ).toBe( basicExpectedQuery );
		//positional args
		var actual = s.csvToQuery( "", path );
		expect( actual ).toBe( basicExpectedQuery ); 
	});

	it( "can read the csv from a text file with an .xls extension", function() {
		var path = ExpandPath( "/root/test/files/csv.xls" );
		var actual = s.csvToQuery( filepath=path );
		expect( actual ).toBe( basicExpectedQuery ); 	
	});

	it( "can accept an alternative delimiter", function() {
		savecontent variable="local.csv"{
			WriteOutput( '
"Frumpo McNugget"|12345
		');
		};
		//named args
		var actual = s.csvToQuery( csv=csv, delimiter="|" );
		expect( actual ).toBe( basicExpectedQuery );
		//positional
		var actual = s.csvToQuery( csv, "", false, true, "|" );
		expect( actual ).toBe( basicExpectedQuery ); 
	});

	it( "can handle an embedded delimiter", function() {
		savecontent variable="local.csv"{
			WriteOutput( '
"McNugget,Frumpo",12345
		');
		};
		var expected = QueryNew( "column1,column2", "", [ [ "McNugget,Frumpo", "12345" ] ] );
		var actual = s.csvToQuery( csv );
		expect( actual ).toBe( expected ); 
	});

	it( "can handle an embedded double-quote", function() {
		savecontent variable="local.csv"{
			WriteOutput( '
"Frumpo ""Frumpie"" McNugget",12345
		');
		};
		var expected = QueryNew( "column1,column2", "", [ [ "Frumpo ""Frumpie"" McNugget", "12345" ] ] );
		var actual = s.csvToQuery( csv );
		expect( actual ).toBe( expected ); 
	});

	it( "can handle an embedded line break", function() {
		savecontent variable="local.csv"{
			WriteOutput( '
"A line#Chr( 10 )#break",12345
		');
		};
		var expected = QueryNew( "column1,column2", "", [ [ "A line#Chr( 10 )#break", "12345" ] ] );
		var actual = s.csvToQuery( csv );
		expect( actual ).toBe( expected ); 
	});

	it( "can handle an embedded line break when there are surrounding spaces", function() {
		savecontent variable="local.csv"{
			WriteOutput( '
A space precedes the next field value, "A line#Chr( 10 )#break"
		');
		};
		var expected = QueryNew( "column1,column2", "", [ [ "A space precedes the next field value", "A line#Chr( 10 )#break" ] ] );
		var actual = s.csvToQuery( csv );
		expect( actual ).toBe( expected ); 
	});

	it( "can handle empty cells", function() {
		savecontent variable="local.csv"{
			WriteOutput( '
Frumpo,McNugget
Susi
Susi,
,Sorglos
		');
		};
		var expected = QueryNew( "column1,column2", "", [ [ "Frumpo", "McNugget" ], [ "Susi", "" ], [ "Susi", "" ], [ "", "Sorglos" ] ] );
		var actual = s.csvToQuery( csv );
		expect( actual ).toBe( expected ); 
	});

	it( "can treat the first line as the column names", function() {
		savecontent variable="local.csv"{
			WriteOutput( '
Name,Phone
Frumpo,12345
		');
		};
		var expected = QueryNew( "Name,Phone", "", [ [ "Frumpo", "12345" ] ] );
		var actual = s.csvToQuery( csv=csv, firstRowIsHeader=true );
		expect( actual ).toBe( expected ); 
	});

	it( "can handle spaces in header/column names", function() {
		savecontent variable="local.csv"{
			WriteOutput( '
Name,Phone Number
Frumpo,12345
		');
		};
		//ACF won't allow spaces in column names when creating queries programmatically. Use Java method to override:
		var expected = QueryNew( "column1,column2", "", [ [ "Frumpo", "12345" ] ] );
		expected.setColumnNames( [ "Name", "Phone Number" ] );
		var actual = s.csvToQuery( csv=csv, firstRowIsHeader=true );
		expect( actual ).toBe( expected ); 
	});

	describe( "throws an exception if", function(){

		it( "neither 'csv' nor 'filepath' are passed", function() {
			expect( function(){
				s.csvToQuery();
			}).toThrow( regex="Missing required argument" );
		});

		it( "both 'csv' and 'filepath' are passed", function() {
			expect( function(){
				s.csvToQuery( csv="x", filepath="x" );
			}).toThrow( regex="Mutually exclusive arguments" );
			expect( function(){
				s.csvToQuery( "x", "x" );
			}).toThrow( regex="Mutually exclusive arguments" );
		});

		it( "a non-existant file is passed", function() {
			expect( function(){
				s.csvToQuery( filepath=ExpandPath( "missing.csv" ) );
			}).toThrow( regex="Non-existant file" );
		});

		it( "a non text/csv file is passed", function() {
			var path = ExpandPath( "/root/test/files/test.xls" );
			expect( function(){
				s.csvToQuery( filepath=path );
			}).toThrow( regex="Invalid csv file" );
		});

	});

});	
</cfscript>