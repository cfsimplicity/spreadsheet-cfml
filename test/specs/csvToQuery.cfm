<cfscript>
describe( "csvToQuery", function(){

	beforeEach( function(){
		variables.basicExpectedQuery = QueryNew( "column1,column2", "", [ [ "Frumpo McNugget", "12345" ] ] );
	});

	it( "converts a basic comma delimited, double quote qualified csv string to a query", function(){
		var csv = '"Frumpo McNugget",12345';
		var actual = s.csvToQuery( csv );
		expect( actual ).toBe( basicExpectedQuery ); 
	});

	it( "can read the csv from a file", function(){
		var path = getTestFilePath( "test.csv" );
		//named args
		var actual = s.csvToQuery( filepath=path );
		expect( actual ).toBe( basicExpectedQuery );
		//positional args
		var actual = s.csvToQuery( "", path );
		expect( actual ).toBe( basicExpectedQuery ); 
	});

	it( "can read the csv from a VFS file", function(){
		var path = "ram:///test.csv";
		if( !DirectoryExists( GetDirectoryFromPath( path ) ) ) //Skip when there's an issue with the ram drive
			return;
		FileCopy( getTestFilePath( "test.csv" ), path );
		var actual = s.csvToQuery( filepath=path );
		expect( actual ).toBe( basicExpectedQuery );
		if( FileExists( path ) )
			FileDelete( path );
	});

	it( "can read the csv from a text file with an .xls extension", function(){
		var path = getTestFilePath( "csv.xls" );
		var actual = s.csvToQuery( filepath=path );
		expect( actual ).toBe( basicExpectedQuery ); 	
	});

	it( "can handle an embedded delimiter", function(){
		var csv = '"McNugget,Frumpo",12345';
		var expected = QueryNew( "column1,column2", "", [ [ "McNugget,Frumpo", "12345" ] ] );
		var actual = s.csvToQuery( csv );
		expect( actual ).toBe( expected ); 
	});

	it( "can handle an embedded double-quote", function(){
		var csv = '"Frumpo ""Frumpie"" McNugget",12345';
		var expected = QueryNew( "column1,column2", "", [ [ "Frumpo ""Frumpie"" McNugget", "12345" ] ] );
		var actual = s.csvToQuery( csv );
		expect( actual ).toBe( expected ); 
	});

	it( "can handle an embedded line break", function(){
		var csv = '"A line#Chr( 10 )#break",12345';
		var expected = QueryNew( "column1,column2", "", [ [ "A line#Chr( 10 )#break", "12345" ] ] );
		var actual = s.csvToQuery( csv );
		expect( actual ).toBe( expected ); 
	});

	it( "can handle an embedded line break when there are surrounding spaces", function(){
		var csv = 'A space precedes the next field value, "A line#Chr( 10 )#break"';
		var expected = QueryNew( "column1,column2", "", [ [ "A space precedes the next field value", "A line#Chr( 10 )#break" ] ] );
		var actual = s.csvToQuery( csv );
		expect( actual ).toBe( expected ); 
	});

	it( "can handle empty cells", function(){
		var csv = 'Frumpo,McNugget#newline#Susi#newline#Susi,#newline#,Sorglos#newline#		';
		var expected = QueryNew( "column1,column2", "", [ [ "Frumpo", "McNugget" ], [ "Susi", "" ], [ "Susi", "" ], [ "", "Sorglos" ] ] );
		var actual = s.csvToQuery( csv );
		expect( actual ).toBe( expected ); 
	});

	it( "can treat the first line as the column names", function(){
		var csv = 'Name,Phone#newline#Frumpo,12345';
		var expected = QueryNew( "Name,Phone", "", [ [ "Frumpo", "12345" ] ] );
		var actual = s.csvToQuery( csv=csv, firstRowIsHeader=true );
		expect( actual ).toBe( expected ); 
	});

	it( "can handle spaces in header/column names", function(){
		var csv = 'Name,Phone Number#newline#Frumpo,12345';
		if( s.getIsACF() ){
			//ACF won't allow spaces in column names when creating queries programmatically. Use setColumnNames() to override:
			var expected = QueryNew( "column1,column2", "", [ [ "Frumpo", "12345" ] ] );
			expected.setColumnNames( [ "Name", "Phone Number" ] );
		}
		else
			var expected = QueryNew( "Name,Phone Number", "", [ [ "Frumpo", "12345" ] ] );
		var actual = s.csvToQuery( csv=csv, firstRowIsHeader=true );
		expect( actual ).toBe( expected ); 
	});

	it( "will preserve the case of header/column names", function(){
		var csv = 'Name,Phone#newline#Frumpo McNugget,12345';
		var actual = s.csvToQuery( csv=csv, firstRowIsHeader=true );
		expect( actual.getColumnNames()[ 1 ] ).toBeWithCase( "Name" );
		//invalid variable name
		csv = '1st Name,Phone#newline#Frumpo McNugget,12345';
		actual = s.csvToQuery( csv=csv, firstRowIsHeader=true );
		expect( actual.getColumnNames()[ 1 ] ).toBeWithCase( "1st Name" );
	});

	describe( "trimming", function(){

		it( "will trim the csv string by default", function(){
			var csv = newline & '"Frumpo McNugget",12345' & newline;
			var actual = s.csvToQuery( csv );
			expect( actual ).toBe( basicExpectedQuery ); 
		});

		it( "will trim the csv file by default", function(){
			var csv = newline & '"Frumpo McNugget",12345' & newline;
			FileWrite( tempCsvPath, csv );
			var actual = s.csvToQuery( filepath: tempCsvPath );
			expect( actual ).toBe( basicExpectedQuery ); 
		});

		it( "can preserve a string's leading/trailing space", function(){
			var csv = newline & '"Frumpo McNugget",12345' & newline;
			var actual = s.csvToQuery( csv: csv, trim: false );
			expected = QueryNew( "column1,column2", "", [ [ "", "" ], [ "Frumpo McNugget", "12345" ] ] );
			expect( actual ).toBe( expected ); 
		});

		it( "can preserve a file's leading/trailing space", function(){
			var csv = newline & '"Frumpo McNugget",12345' & newline;
			FileWrite( tempCsvPath, csv );
			var actual = s.csvToQuery( filepath: tempCsvPath, trim: false );
			expected = QueryNew( "column1,column2", "", [ [ "", "" ], [ "Frumpo McNugget", "12345" ] ] );
			expect( actual ).toBe( expected ); 
		});

		afterEach( function(){
			if( FileExists( tempCsvPath ) )
				FileDelete( tempCsvPath );
		});

	});

	describe( "delimiter handling", function(){

		it( "can accept an alternative delimiter", function(){
			var csv = '"Frumpo McNugget"|12345';
			//named args
			var actual = s.csvToQuery( csv=csv, delimiter="|" );
			expect( actual ).toBe( basicExpectedQuery );
			//positional
			var actual = s.csvToQuery( csv, "", false, true, "|" );
			expect( actual ).toBe( basicExpectedQuery ); 
		});

		it( "has special handling for tab delimited data", function(){
			var csv = '"Frumpo McNugget"#Chr( 9 )#12345';
			var validTabValues = [ "#Chr( 9 )#", "\t", "tab", "TAB" ];
			for( var value in validTabValues ){
				var actual = s.csvToQuery( csv=csv, delimiter="#value#" );
				expect( actual ).toBe( basicExpectedQuery );
			}
		});

	});

	describe( "query column name setting", function(){

		it( "Allows column names to be specified as an array when reading a csv into a query", function(){
			var csv = '"Frumpo McNugget",12345';
			var columnNames = [ "name", "phone number" ];
			var q = s.csvToQuery( csv=csv, queryColumnNames=columnNames, firstRowIsHeader=true );
			expect( q.getColumnNames()[ 1 ] ).toBe( columnNames[ 1 ] );
			expect( q.getColumnNames()[ 2 ] ).toBe( columnNames[ 2 ] );
		});

		it( "ColumnNames argument overrides firstRowIsHeader: none of the header row values will be used", function(){
			var csv = 'header1,header2#newline#"Frumpo McNugget",12345';
			var columnNames = [ "name", "phone number" ];
			var q = s.csvToQuery( csv=csv, queryColumnNames=columnNames );
			expect( q.getColumnNames()[ 1 ] ).toBe( columnNames[ 1 ] );
			expect( q.getColumnNames()[ 2 ] ).toBe( columnNames[ 2 ] );
		});

		it( "Allows csv header names to be made safe for query column names", function(){
			var csv = 'id,id,"A  B","x/?y","(a)"," A","##1","1a"#newline#1,2,3,4,5,6,7,8';
			var q = s.csvToQuery( csv=csv, firstRowIsHeader=true, makeColumnNamesSafe=true );
			expect( q.getColumnNames() ).toBe( [ "id", "id2", "A_B", "x_y", "_a_", "A", "Number1", "_a" ] );
		});

	});

	describe( "query column type setting", function(){

		it( "allows the query column types to be manually set using a list", function(){
			var csv = '1,1.1,"string",#CreateTime( 1, 0, 0 )#';
			var q = s.csvToQuery( csv=csv, queryColumnTypes="Integer,Double,VarChar,Time" );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "INTEGER" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIME" );
		});

		it( "allows the query column types to be manually set where the column order isn't known, but the header row values are", function(){
			var csv = 'integer,double,"string column",time#newline#1,1.1,string,12:00';
			var columnTypes = { "string column": "VARCHAR", "integer": "INTEGER", "time": "TIME", "double": "DOUBLE" };//not in order
			var q = s.csvToQuery( csv=csv, queryColumnTypes="Integer,Double,VarChar,Time", firstRowIsHeader=true );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "INTEGER" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIME" );
		});

		it( "allows the query column types to be manually set where the column order isn't known, but the column names are", function(){
			var csv = '1,1.1,"string",#CreateTime( 1, 0, 0 )#';
			var columnNames = [ "integer", "double", "string column", "time" ];
			var columnTypes = { "string": "VARCHAR", "integer": "INTEGER", "time": "TIME", "double": "DOUBLE" };//not in order
			var q = s.csvToQuery( csv=csv, queryColumnTypes=columnTypes, queryColumnNames=columnNames );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "INTEGER" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIME" );
		});

		it( "allows the query column types to be automatically set", function(){
			var csv = '1,1.1,"string",2021-03-10 12:00:00';
			var q = s.csvToQuery( csv=csv, queryColumnTypes="auto" );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIMESTAMP" );
		});

		it( "automatic detecting of query column types ignores blank cells", function(){
			var csv = ',,,#newline#,2,test,2021-03-10 12:00:00#newline#1,1.1,string,2021-03-10 12:00:00#newline#1,,,';
			var q = s.csvToQuery( csv=csv, queryColumnTypes="auto" );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "TIMESTAMP" );
		});

		it( "allows a default type to be set for all query columns", function(){
			var csv = '1,1.1,"string",#CreateTime( 1, 0, 0 )#';
			var q = s.csvToQuery( csv=csv, queryColumnTypes="VARCHAR" );
			var columns = GetMetaData( q );
			expect( columns[ 1 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 2 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
			expect( columns[ 4 ].typeName ).toBe( "VARCHAR" );
		});

	});

	describe( "csvToQuery throws an exception if", function(){

		it( "neither 'csv' nor 'filepath' are passed", function(){
			expect( function(){
				s.csvToQuery();
			}).toThrow( type="cfsimplicity.spreadsheet.missingRequiredArgument" );
		});

		it( "both 'csv' and 'filepath' are passed", function(){
			expect( function(){
				s.csvToQuery( csv="x", filepath="x" );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidArgumentCombination" );
			expect( function(){
				s.csvToQuery( "x", "x" );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidArgumentCombination" );
		});

		it( "a non-existent file is passed", function(){
			expect( function(){
				s.csvToQuery( filepath=ExpandPath( "missing.csv" ) );
			}).toThrow( type="cfsimplicity.spreadsheet.nonExistentFile" );
		});

		it( "a non text/csv file is passed", function(){
			var path = getTestFilePath( "test.xls" );
			expect( function(){
				s.csvToQuery( filepath=path );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidCsvFile" );
		});

		it( "queryColumnTypes is specified as a 'columnName/type' struct, but firstRowIsHeader is not set to true AND columnNames are not provided", function(){
			expect( function(){
				// using 'var' keyword here causes ACF2021 to throw exception
				local.columnTypes = { col1: "Integer" };
				local.q = s.csvToQuery( csv="1", queryColumnTypes=columnTypes );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidArgumentCombination" );
		});

	});

});	
</cfscript>