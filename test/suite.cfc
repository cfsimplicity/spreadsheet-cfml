component extends="testbox.system.BaseSpec"{

	function beforeAll(){
		variables.tempXlsPath = ExpandPath( "temp.xls" );
	}

	function afterAll(){}

	function run( testResults,testBox ){

		describe( "spreadsheet test suite",function() {
     
			beforeEach( function( currentSpec ) {
			  variables.s = New root.spreadsheet();
			  makePublic( s,"sheetToQuery" );
			});

			afterEach(function( currentSpec ) {
		    if( FileExists( tempXlsPath ) )
					FileDelete( tempXlsPath );
			});

			var testFiles = DirectoryList( ExpandPath( "tests" ),false,"name","*.cfm" );
			// run every file in the tests folder
			for( var file in testFiles ){
				include "tests/#file#";	
			}

		});

	}

	//dump( expected );dump( actual );abort;

}
