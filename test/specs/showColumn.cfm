<cfscript>
describe( "showColumn", function(){

	beforeEach( function(){
		var query = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a","b" ], [ "c","d" ] ] );
		var xls = s.workbookFromQuery( query );
		var xlsx = s.workbookFromQuery( data=query, xmlFormat=true );
		variables.workbooks = [ xls, xlsx ];
	});

	it( "can show a column", function(){
		workbooks.Each( function( wb ){
			s.hideColumn( wb, 1 );
			expect( s.isColumnHidden( wb, 1 ) ).toBeTrue();
			s.showColumn( wb, 1 );
			expect( s.isColumnHidden( wb, 1 ) ).toBeFalse();
		});
	});

	it( "is chainable", function(){
		workbooks.Each( function( wb ){
			s.newChainable( wb ).hideColumn( 1 );
			expect( s.isColumnHidden( wb, 1 ) ).toBeTrue();
			s.newChainable( wb ).showColumn( 1 );
			expect( s.isColumnHidden( wb, 1 ) ).toBeFalse();
		});
	});

});	
</cfscript>