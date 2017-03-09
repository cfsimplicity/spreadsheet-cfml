component extends="testbox.system.BaseSpec"{

	function beforeAll(){
		variables.s = New root.Spreadsheet();
	  makePublic( s, "sheetToQuery" );
	  variables.tempXlsPath = ExpandPath( "temp.xls" );
	  variables.tempXlsxPath = ExpandPath( "temp.xlsx" );
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