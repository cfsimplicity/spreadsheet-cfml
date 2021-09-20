<cfscript>
describe( "cellComment", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
	});

	it( "Can set and get a comment from the specified cell", function() {
		var theComment = {
			author: "cfsimplicity"
			,comment: "This is the comment in row 1 column 1"
		};
		var expected = Duplicate( theComment ).Append( { column: 1, row: 1 } );
		workbooks.Each( function( wb ){
			s.addColumn( wb, "1" )
				.setCellComment( wb, theComment, 1, 1 );
			var actual = s.getCellComment( wb, 1, 1 );
			expect( actual ).toBe( expected );
		});
	});

	it( "getCellComment, getCellComments and setCellComment are chainable", function() {
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
		var dataAsArray = [ [ "a", "b" ], [ "c", "d" ] ];
		workbooks.Each( function( wb ){
			s.addRows( wb, data );
			var comments = [];
			comments.Append( { author: "cfsimplicity", comment: "This is the comment in row 1 column 1", column: 1, row: 1 } );
			comments.Append( { author: "cfsimplicity", comment: "This is the comment in row 2 column 2", column: 2, row: 2 } );
			var wbChainable = s.newChainable( wb )
				.setCellComment( comments[ 1 ], 1, 1 )
				.setCellComment( comments[ 2 ], 2, 2 );
			expect( wbChainable.getCellComment() ).toBe( comments );
			expect( wbChainable.getCellComments() ).toBe( comments );
			expect( wbChainable.getCellComment( 1, 1 ) ).toBe( comments[ 1 ] );
		});
	});

	it( "Can get all comments in the current sheet", function() {
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
		var dataAsArray = [ [ "a", "b" ], [ "c", "d" ] ];
		workbooks.Each( function( wb ){
			s.addRows( wb, data );
			var comments = [];
			comments.Append( { author: "cfsimplicity", comment: "This is the comment in row 1 column 1", column: 1, row: 1 } );
			comments.Append( { author: "cfsimplicity", comment: "This is the comment in row 2 column 2", column: 2, row: 2 } );
			s.setCellComment( wb, comments[ 1 ], 1, 1 )
				.setCellComment( wb, comments[ 2 ], 2, 2 );
			var expected = comments;
			var actual = s.getCellComment( wb );
			expect( actual ).toBe( expected );
			//alias getCellComments
			actual = s.getCellComments( wb );
			expect( actual ).toBe( expected );
		});
	});

	it( "can set comment styles without erroring", function(){
		workbooks.Each( function( wb ){
			s.addColumn( wb, "1" );
			var theComment = {
				anchor: "1,2,3,4"
				,author: "cfsimplicity"
				,bold: "true"
				,comment: "This is the comment in row 1 column 1"
				,color: "blue"
				,font: "Times New Roman"
				,italic: "true"
				,size: 16
				,strikeout: "true"
				,underline: "true"
				,visible: "true"
				//following 5 not supported by xlsx
				,fillcolor: "magenta"
				,horizontalalignment: "center"
				,linestyle: "dashsys"
				,linestylecolor: "cyan"
				,verticalalignment: "center"
			};
			s.setCellComment( wb, theComment, 1, 1 );
		});
	});

	describe( "cellComment throws an exception if", function(){

		it( "column specified but not row, or vice versa", function() {
			workbooks.Each( function( wb ){
				expect( function(){
					s.getCellComment( workbook=wb, row=1 );
					s.getCellComment( workbook=wb, column=1 );
				}).toThrow( regex="Invalid argument combination" );
			});
		});

	});	

});	
</cfscript>