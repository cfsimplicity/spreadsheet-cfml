<cfscript>
describe( "getEnvironment", function(){

	it( "returns a struct with the expected keys", function() {
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
	});

});	
</cfscript>