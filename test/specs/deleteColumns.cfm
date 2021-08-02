<cfscript>
describe( "deleteColumns", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Deletes the data in a specified range of columns", function(){
		var expected = querySim("column1,column2,column3,column4,column5
			||e||i
			||f||j");
		workbooks.Each( function( wb ){
			s.addColumn( wb, "a,b" )
				.addColumn( wb, "c,d" )
				.addColumn( wb, "e,f" )
				.addColumn( wb, "g,h" )
				.addColumn( wb, "i,j" )
				.deleteColumns( wb, "1-2,4" );
			var actual = s.getSheetHelper().sheetToQuery( workbook=wb, includeBlankRows=true );
			expect( actual ).toBe( expected );
		});
	});

	describe( "deleteColumns throws an exception if", function(){

		it( "the range is invalid", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.deleteColumns( wb, "a-b" );
				}).toThrow( regex="Invalid range" );
			});
		});

	});

});	
</cfscript>