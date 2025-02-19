<cfscript>
describe( "getEnvironment", ()=>{

	it( "returns a struct with the expected keys", ()=>{
		var expectedKeys = [
			"dateFormats"
			,"engine"
			,"javaLoaderDotPath"
			,"javaClassesLastLoadedVia"
			,"javaLoaderName"
			,"javaVersion"
			,"requiresJavaLoader"
			,"version"
			,"poiVersion"
			,"osgiLibBundleVersion"
		];
		var actual = s.getEnvironment();
		for( var key in expectedKeys ){
			expect( actual ).toHaveKey( key );
		}
	})

})
</cfscript>