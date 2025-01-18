<cfscript>
describe( "java class loading", function(){

	it( "defaults to the appropriate method for the engine", function() {
		s.getPoiVersion();
		//defaults
		if( s.getIsACF() ){
			expect( s.getLoadJavaClassesUsing() ).toBe( "JavaLoader" );
			expect( s.getJavaClassesLastLoadedVia() ).toBe( "JavaLoader" );
		}
		if( s.getIsLucee() ){
			expect( s.getLoadJavaClassesUsing() ).toBe( "osgi" );
			expect( s.getJavaClassesLastLoadedVia() ).toBe( "OSGi bundle" );
		}
	});

	it( "can be configured at instance level", function() {
		local.s = newSpreadsheetInstance( loadJavaClassesUsing="dynamicPath" );
		expect( local.s.getLoadJavaClassesUsing() ).toBe( "dynamicPath" );
		if( s.getIsLucee() ){
			//default is OSGi. Let's override this
			local.s = newSpreadsheetInstance();
			local.s.setLoadJavaClassesUsing( "JavaLoader" );
			local.s.getPoiVersion();
			expect( local.s.getJavaClassesLastLoadedVia() ).toBe( "JavaLoader" );
		}
	});

	it( "throws an exception if an invalid loading method is specified", function() {
		expect( function(){
			local.s = newSpreadsheetInstance( loadJavaClassesUsing="invalid" );
		}).toThrow( type="cfsimplicity.spreadsheet.invalidJavaLoadingMethod" );
	});

});	
</cfscript>