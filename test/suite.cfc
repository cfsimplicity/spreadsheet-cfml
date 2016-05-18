component extends="testbox.system.BaseSpec"{

	function beforeAll(){}

	function afterAll(){}

	function run( testResults,testBox ){

		describe( "spreadsheet test suite",function() {
     
			beforeEach( function( currentSpec ) {
			  variables.s=New root.Spreadsheet();
			  makePublic( s,"sheetToQuery" );
			  variables.tempXlsPath=ExpandPath( "temp.xls" );
			});

			afterEach(function( currentSpec ) {
		    if( FileExists( tempXlsPath ) )
					FileDelete( tempXlsPath );
			});

			var specs=DirectoryList( ExpandPath( "specs" ),false,"name","*.cfm" );
			// run every file in the tests folder
			for( var file in specs ){
				include "specs/#file#";	
			}

		});

	}

	//dump( expected );dump( actual );abort;

}