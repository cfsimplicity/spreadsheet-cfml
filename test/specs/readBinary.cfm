<cfscript>
describe( "readBinary", ()=>{

	it( "Returns a binary object", ()=>{
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( ( wb )=>{
			expect( IsBinary( s.readBinary( wb ) ) ).toBeTrue();
		})
	})

	it( "Is chainable", ()=>{
		var workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( ( wb )=>{
			var actual = s.newChainable( wb ).readBinary();
			expect( IsBinary( actual ) ).toBeTrue();
		})
	})

})	
</cfscript>