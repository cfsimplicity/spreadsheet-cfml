<cfscript>
describe( "formatRows", function(){

	describe( "formatRows throws an exception if", function(){

		it( "the range is invalid", function(){
			var workbooks = [ s.newXls(), s.newXlsx() ];
			workbooks.Each( function( wb ){
				expect( function(){
					var format = { font: "Courier" };
					s.formatRows( wb, format, "a-b" );
				}).toThrow( regex="Invalid range" );
			});
		});

	});

});	
</cfscript>