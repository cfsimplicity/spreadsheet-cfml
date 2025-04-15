<cfscript>
describe( "java class loading", ()=>{

	it( "defaults to the appropriate method for the engine", ()=>{
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
	})

	it( "can be configured at instance level", ()=>{
		local.s = newSpreadsheetInstance( loadJavaClassesUsing="dynamicPath" );
		expect( local.s.getLoadJavaClassesUsing() ).toBe( "dynamicPath" );
	})

	it( "throws an exception if an invalid loading method is specified", ()=>{
		expect( ()=>{
			local.s = newSpreadsheetInstance( loadJavaClassesUsing="invalid" );
		}).toThrow( type="cfsimplicity.spreadsheet.invalidJavaLoadingMethod" );
	})

})	
</cfscript>