<cfscript>
describe( "createJavaObject", ()=>{

	it( "creates a java object from the bundled library jars", ()=>{
		var className = "org.apache.poi.Version";
		var object = s.createJavaObject( className );
		expect( object.getClass().getCanonicalName() ).toBe( className );
	})

})	
</cfscript>