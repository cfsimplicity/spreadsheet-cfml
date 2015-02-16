component extends="testbox.system.BaseSpec"{

	function beforeAll(){
		variables.tempXlsPath = ExpandPath( "temp.xls" );
	}

	function afterAll(){}

	/* helpers */
	query function data(){
		return QueryNew( "First,Last","VarChar,VarChar",[ [ "Susi","Sorglos" ],[ "Julian","Halliwell" ] ] );
	}

	function run( testResults,testBox ){

		describe( "spreadsheet test suite",function() {
     
			beforeEach( function( currentSpec ) {
			  variables.s = New root.spreadsheet();
			});

			afterEach(function( currentSpec ) {
		    if( FileExists( tempXlsPath ) )
					FileDelete( tempXlsPath );
			});

			include "test/addColumn.cfm";

		});

	}

}