<cfscript>
describe( "cellComment", function(){

	it( "Can set and get a comment from the specified cell", function() {
		var xls = s.new();
		s.addColumn( xls, "1" );
		var theComment = {
			author="cfsimplicity"
			,comment="This is the comment in row 1 column 1"
		};
		var expected = Duplicate( theComment ).Append( { column: 1, row: 1 } );
		s.setCellComment( xls, theComment, 1, 1 );
		var actual = s.getCellComment( xls, 1, 1 );
		expect( actual ).toBe( expected );
		//xlsx
		var xlsx = s.newXlsx();
		s.addColumn( xlsx, "1" );
		s.setCellComment( xlsx, theComment, 1, 1 );
		actual = s.getCellComment( xlsx, 1, 1 );
		expect( actual ).toBe( expected );
	});

	it( "Can get all comments in the current sheet", function() {
		var data = QueryNew( "column1,column2", "VarChar,VarChar", [ [ "a", "b" ], [ "c", "d" ] ] );
		var dataAsArray = [ [ "a", "b" ], [ "c", "d" ] ];
		var workbook = s.newXls();
		s.addRows( workbook, data );
		var comments = [];
		comments.Append( { author: "cfsimplicity", comment: "This is the comment in row 1 column 1" } );
		comments.Append( { author: "cfsimplicity", comment: "This is the comment in row 2 column 2" } );
		s.setCellComment( workbook, comments[ 1 ], 1, 1 );
		s.setCellComment( workbook, comments[ 2 ], 2, 2 );
		comments[ 1 ].Append( { column: 1, row: 1 } );
		comments[ 2 ].Append( { column: 2, row: 2 } );
		var expected = comments;
		var actual = s.getCellComment( workbook );
		expect( actual ).toBe( expected );
		//alias getCellComments
		actual = s.getCellComments( workbook );
		expect( actual ).toBe( expected );
	});

	it( "can set comment styles without erroring", function(){
		var xls = s.new();
		s.addColumn( xls, "1" );
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
		s.setCellComment( xls, theComment, 1, 1 );
		var xlsx = s.newXlsx();
		s.addColumn( xlsx, "1" );
		s.setCellComment( xlsx, theComment, 1, 1 );
	});

	describe( "cellComment throws an exception if", function(){

		it( "column specified but not row, or vice versa", function() {
			expect( function(){
				s.getCellComment( workbook=workbook, row=1 );
				s.getCellComment( workbook=workbook, column=1 );
			}).toThrow( regex="Invalid argument combination" );
		});

	});	

});	
</cfscript>