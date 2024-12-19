component extends="testbox.system.BaseSpec"{

	function newSpreadsheetInstance(){
		var s = New root.Spreadsheet( argumentCollection=arguments );
		return s;
	}

	//Allow universal access including outside tests
	variables.s = newSpreadsheetInstance();
	
	function beforeAll(){
		s.flushOsgiBundle();
		if( server.KeyExists( s.getJavaLoaderName() ) )
			server.delete( s.getJavaLoaderName() );
	  variables.tempXlsPath = ExpandPath( "temp.xls" );
	  variables.tempXlsxPath = ExpandPath( "temp.xlsx" );
	  variables.tempCsvPath = ExpandPath( "temp.csv" );
	  variables.newline = Chr( 13 ) & Chr( 10 );
	  variables.spreadsheetTypes = [ "xls", "xlsx" ];
	}

	function getTestFilePath( required string filename ){
		return ExpandPath( "/root/test/files/" ) & arguments.filename;
	}

	function afterAll(){
		WriteDump( var=s.getEnvironment(), label="Environment and settings" );
	}

	function run( testResults, testBox ){

		var specs = DirectoryList( ExpandPath( "specs" ), false, "name", "*.cfm" );
		// run every file in the tests folder
		for( var file in specs ){
			include "specs/#file#";
		}

	}

}