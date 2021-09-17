<cfscript>
describe( "addAutoFilter", function(){

	beforeEach( function(){
		var data = QueryNew( "Header1,Header2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
		var xls = s.workbookFromQuery( data );
		var xlsx = s.workbookFromQuery( data=data, xmlformat=true );
		variables.workbooks = [ xls, xlsx ];
	});

	it( "Doesn't error when passing valid arguments", function() {
		workbooks.Each( function( wb ){
			s.addAutoFilter( wb, "A1:B1" )
			.addAutoFilter( wb )// default to all cols in first row if no row range passed
			.addAutoFilter( workbook=wb, row=2 );// allow row to be specified instead of range
		});
		
	});

	it( "Doesn't error when passing valid arguments with extra trailing/leading space", function() {
		workbooks.Each( function( wb ){
			s.addAutoFilter( wb, " A1:B1 " );
		});
	});

	it( "Is chainable", function() {
		workbooks.Each( function( wb ){
			s.newChainable( wb ).addAutoFilter( " A1:B1 " );
		});
	});

});	
</cfscript>