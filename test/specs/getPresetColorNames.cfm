<cfscript>
describe( "getPresetColorNames", ()=>{

	it( "returns an alphabetical array of preset color names available for use in color formatting options", ()=>{
		expect( s.getPresetColorNames() ).toHaveLength( 48 );
		expect( s.getPresetColorNames()[ 1 ] ).toBe( "AQUA" );
		expect( s.getPresetColorNames()[ 48 ] ).toBe( "YELLOW" );
	})

});
</cfscript>