component extends="testbox.system.BaseSpec"{

	function newSpreadsheetInstance(){
		var s = New root.Spreadsheet( argumentCollection=arguments );
		return s;
	}

	//Allow universal access including outside tests
	variables.s = newSpreadsheetInstance();

	// These are differences between ACF/Lucee and Boxlang engine
	variables.DOUBLE_CHECK = !s.getIsBoxlang() ? "DOUBLE" : "NUMERIC";
	variables.VARCHAR_CHECK = !s.getIsBoxlang() ? "VARCHAR" : "STRING";
	variables.TIME_WITH_MILLISECONDS_MASK = !s.getIsBoxlang() ? "hh:nn:ss:l" : "hh:nn:ss.L";

	// BoxLang's QueryNew stores "" as null for VarChar columns even when explicitly passed.
	// s.read() always returns "" for blank cells. Replace nulls with "" after building expected queries.
	variables.normalizeExpectedQuery = ( required query q ) => {
		if( !s.getIsBoxlang() )
			return q;
		q.Each( ( row, rowNumber ) => {
			row.Each( ( key, value ) => {
				if( IsNull( value ) || !value.Len() )
					QuerySetCell( q, key, "", rowNumber );
			});
		});
		return q;
	};
	
	function beforeAll(){
		s.flushOsgiBundle();
		if( server.KeyExists( s.getJavaLoaderName() ) )
			server.delete( s.getJavaLoaderName() );
	  variables.tempXlsPath = GetTempDirectory() & "temp.xls";
	  variables.tempXlsxPath = GetTempDirectory() & "temp.xlsx";
	  variables.tempCsvPath = GetTempDirectory() & "temp.csv";
	  variables.newline = Chr( 13 ) & Chr( 10 );
	  variables.spreadsheetTypes = [ "xls", "xlsx" ];
	}

	function getTestFilePath( required string filename ){
		return ExpandPath( "/root/test/files/" ) & arguments.filename;
	}

	function _CreateTime( required numeric hours, required numeric minutes, required numeric seconds ){
		//boxlang does not have CreateTime() BIF
		return CreateDateTime( 1899, 12, 30, arguments.hours, arguments.minutes, arguments.seconds );
	}

	function afterAll(){
		if( !url.keyExists( "reporter" ) || url.reporter != "json" )
			WriteDump( var=s.getEnvironment(), label="Environment and settings" );
		if( FileExists( variables.tempXlsPath ) )
			FileDelete( variables.tempXlsPath );
		if( FileExists( variables.tempXlsxPath ) )
			FileDelete( variables.tempXlsxPath );
		if( FileExists( variables.tempCsvPath ) )
			FileDelete( variables.tempCsvPath );
	}

	function run( testResults, testBox ){

		var specs = DirectoryList( ExpandPath( "specs" ), false, "name", "*.cfm" );
		// run every file in the tests folder
		for( var file in specs ){
			include "specs/#file#";
		}

	}

}