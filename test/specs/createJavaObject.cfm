<cfscript>
describe( "createJavaObject", function(){

	it( "creates a java object from the bundled library jars", function() {
		var className = "org.apache.poi.Version";
		var object = s.createJavaObject( className );
		expect( object.getClass().getCanonicalName() ).toBe( className );
	});

});	
</cfscript>