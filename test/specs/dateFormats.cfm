<cfscript>
describe( "dateFormats customisability",function(){

	it( "the default dateFormats can be overridden individually",function() {
		var s=New root.spreadsheet();
		expected={
			DATE="yyyy-mm-dd"
			,DATETIME="yyyy-mm-dd HH:nn:ss"
			,TIME="hh:mm:ss"
			,TIMESTAMP="yyyy-mm-dd hh:mm:ss"
		};
		actual=s.getDateFormats();
		expect( actual ).toBe( expected );
		s=New root.spreadsheet( dateFormats={ DATE="mm/dd/yyyy" } );
		expected={
			DATE="mm/dd/yyyy"
			,DATETIME="yyyy-mm-dd HH:nn:ss"
			,TIME="hh:mm:ss"
			,TIMESTAMP="yyyy-mm-dd hh:mm:ss"
		};
		actual=s.getDateFormats();
		expect( actual ).toBe( expected );
	});

	it( "Throws an exception if a passed format key is invalid",function() {
		expect( function(){
			s=New root.spreadsheet( dateFormats={ DAT="mm/dd/yyyy" } );
		}).toThrow( regex="Invalid date format key" );
	});

});	
</cfscript>