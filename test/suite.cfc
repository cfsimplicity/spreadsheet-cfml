component extends="testbox.system.BaseSpec"{

	//Allow universal access including outside tests
	variables.s = New root.Spreadsheet();

	function beforeAll(){
	  makePublic( s, "sheetToQuery" );
	  variables.tempXlsPath = ExpandPath( "temp.xls" );
	  variables.tempXlsxPath = ExpandPath( "temp.xlsx" );
	  WriteDump( var=s.getEnvironment(), label="Environment and settings" );
	}

	function afterAll(){}

	function run( testResults,testBox ){

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