component extends="testbox.system.BaseSpec"{

	function newSpreadsheetInstance(){
		var s = New root.Spreadsheet( argumentCollection=arguments );
		makePublic( s, "sheetToQuery" );
		return s;
	}

	//Allow universal access including outside tests
	variables.s = newSpreadsheetInstance();
	
	function beforeAll(){
	  variables.filesDirectoryPath = ExpandPath( "files/" );
	  variables.tempXlsPath = ExpandPath( "temp.xls" );
	  variables.tempXlsxPath = ExpandPath( "temp.xlsx" );
	}

	function getTestFilePath( required string filename ){
		return filesDirectoryPath & filename;
	}

	function afterAll(){
		WriteDump( var=s.getEnvironment(), label="Environment and settings" );
	}

	function run( testResults, testBox ){

		describe( "spreadsheet test suite",function() {
     
			/* beforeEach( function( currentSpec ) {}); */

			afterEach(function( currentSpec ) {
		    if( FileExists( tempXlsPath ) )
		    	FileDelete( tempXlsPath );
		    if( FileExists( tempXlsxPath ) )
		    	FileDelete( tempXlsxPath );
			});

			var specs = DirectoryList( ExpandPath( "specs" ), false, "name", "*.cfm" );
			// run every file in the tests folder
			for( var file in specs )
				include "specs/#file#";	

		});

	}

}