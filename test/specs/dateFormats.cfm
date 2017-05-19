<cfscript>
describe( "dateFormats customisability",function(){

	it( "the default dateFormats can be overridden individually",function() {
		var s=newSpreadsheetInstance();
		expected={
			DATE="yyyy-mm-dd"
			,DATETIME="yyyy-mm-dd HH:nn:ss"
			,TIME="hh:mm:ss"
			,TIMESTAMP="yyyy-mm-dd hh:mm:ss"
		};
		actual=s.getDateFormats();
		expect( actual ).toBe( expected );
		s=newSpreadsheetInstance( dateFormats={ DATE="mm/dd/yyyy" } );
		expected={
			DATE="mm/dd/yyyy"
			,DATETIME="yyyy-mm-dd HH:nn:ss"
			,TIME="hh:mm:ss"
			,TIMESTAMP="yyyy-mm-dd hh:mm:ss"
		};
		actual=s.getDateFormats();
		expect( actual ).toBe( expected );
	});

	describe( "dateFormats: throws an exception if",function(){

		it( "a passed format key is invalid",function() {
			expect( function(){
				var s=newSpreadsheetInstance( dateFormats={ DAT="mm/dd/yyyy" } );
			}).toThrow( regex="Invalid date format key" );
		});

	});	


});	
</cfscript>