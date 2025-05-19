<cfscript>
describe( "read: format=query", ()=>{

  describe( "query column name setting", ()=>{

		it( "Allows header names to be made safe for query column names", ()=>{
			var data = [ [ "id","id","A  B","x/?y","(a)"," A","##1","1a" ], [ 1,2,3,4,5,6,7,8 ] ];
			variables.spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRows( data )
					.write( path, true );
				var q = s.read( src=path, format="query", headerRow=1, makeColumnNamesSafe=true );
				var expected = [ "id", "id2", "A_B", "x_y", "_a_", "A", "Number1", "_a" ];
				cfloop( from=1, to=expected.Len(), index="i" ){
					expect( q.getColumnNames()[ i ] ).toBe( expected[ i ] );
				}
			})
		})

	})

  describe( "query column type setting", ()=>{

		it( "allows the query column types to be manually set using a list", ()=>{
			var data = [ [ 1, 1.1, "string", _CreateTime( 1, 0, 0 ) ] ]
			variables.spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRows( data )
					.write( path, true )
				var q = s.read( src=path, format="query", queryColumnTypes="Integer,Double,VarChar,Time" );
				var columns = GetMetaData( q );
				expect( columns[ 1 ].typeName ).toBe( "INTEGER" );
				expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
				expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
				expect( columns[ 4 ].typeName ).toBe( "TIME" );
			})
		})

		it( "allows the query column types to be manually set where the column order isn't known, but the header row values are", ()=>{
			var data = [ [ "integer", "double", "string column", "time" ], [ 1, 1.1, "text", _CreateTime( 1, 0, 0 ) ] ];
			variables.spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRows( data )
					.write( path, true );
				var columnTypes = { "string column": "VARCHAR", "integer": "INTEGER", "time": "TIME", "double": "DOUBLE" };//not in order
				var q = s.read( src=path, format="query", queryColumnTypes=columnTypes, headerRow=1 );
				var columns = GetMetaData( q );
				expect( columns[ 1 ].typeName ).toBe( "INTEGER" );
				expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
				expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
				expect( columns[ 4 ].typeName ).toBe( "TIME" );
			})
		})

		it( "allows the query column types to be manually set where the column order isn't known, but the column names are", ()=>{
			var data = [ [ 1, 1.1, "text", _CreateTime( 1, 0, 0 ) ] ];
			variables.spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRows( data )
					.write( path, true );
				var columnNames = "integer,double,string column,time";
				var columnTypes = { "string": "VARCHAR", "integer": "INTEGER", "time": "TIME", "double": "DOUBLE" };//not in order
				var q = s.read( src=path, format="query", queryColumnTypes=columnTypes, columnNames=columnNames );
				var columns = GetMetaData( q );
				expect( columns[ 1 ].typeName ).toBe( "INTEGER" );
				expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
				expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
				expect( columns[ 4 ].typeName ).toBe( "TIME" );
			})
		})

		it( "allows the query column types to be automatically set", ()=>{
			var data = [ [ 1, 1.1, "string", Now() ] ];
			variables.spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRows( data )
					.write( path, true );
				var q = s.read( src=path, format="query", queryColumnTypes="auto" );
				var columns = GetMetaData( q );
				expect( columns[ 1 ].typeName ).toBe( "DOUBLE" );
				expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
				expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
				expect( columns[ 4 ].typeName ).toBe( "TIMESTAMP" );
			})
		})

		it( "automatic detecting of query column types ignores blank cells", ()=>{
			var data = [
				[ "", "", "", "" ],
				[ "", 2, "test", Now() ],
				[ 1, 1.1, "string", Now() ],
				[ 1, "", "", "" ]
			];
			variables.spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRows( data )
					.write( path, true );
				var q = s.read( src=path, format="query", queryColumnTypes="auto" );
				var columns = GetMetaData( q );
				expect( columns[ 1 ].typeName ).toBe( "DOUBLE" );
				expect( columns[ 2 ].typeName ).toBe( "DOUBLE" );
				expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
				expect( columns[ 4 ].typeName ).toBe( "TIMESTAMP" );
			})
		})

		it( "allows a default type to be set for all query columns", ()=>{
			var data = [ [ 1, 1.1, "string", Now() ] ];
			variables.spreadsheetTypes.Each( ( type )=>{
				var path = variables[ "temp" & type & "Path" ];
				s.newChainable( type )
					.addRows( data )
					.write( path, true );
				var q = s.read( src=path, format="query", queryColumnTypes="VARCHAR" );
				var columns = GetMetaData( q );
				expect( columns[ 1 ].typeName ).toBe( "VARCHAR" );
				expect( columns[ 2 ].typeName ).toBe( "VARCHAR" );
				expect( columns[ 3 ].typeName ).toBe( "VARCHAR" );
				expect( columns[ 4 ].typeName ).toBe( "VARCHAR" );
			})
		})

	})

})
</cfscript>