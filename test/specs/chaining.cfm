<cfscript>
describe( "chaining", function(){

	it( "Allows void methods to be chained", function() {
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		var theComment = {
			author: "cfsimplicity"
			,comment: "This is the comment in row 1 column 1"
		};
		var expected = Duplicate( theComment ).Append( { column: 1, row: 1 } );
		workbooks.Each( function( wb ){
			s.addColumn( wb, "1" ).setCellComment( wb, theComment, 1, 1 );
			var actual = s.getCellComment( wb, 1, 1 );
			expect( actual ).toBe( expected );
		});
	});

	describe( "newChainable", function(){

		describe( "initialisation", function() {

			it( "allows a new workbook of the specified type to be created inside a chainable object", function(){
				o = s.newChainable( "xls" );
				expect( s.isBinaryFormat( o.getWorkbook() ) ).toBeTrue();
				o = s.newChainable( "xlsx" );
				expect( s.isXmlFormat( o.getWorkbook() ) ).toBeTrue();
				o = s.newChainable( "streamingXml" );
				expect( s.isStreamingXmlFormat( o.getWorkbook() ) ).toBeTrue();
				o = s.newChainable( "streamingXlsx" );
				expect( s.isStreamingXmlFormat( o.getWorkbook() ) ).toBeTrue();
			});

			it( "allows an existing wookbook to be passed to a chainable object", function(){
				xls = s.newXls();
				o = s.newChainable( xls );
				expect( s.isBinaryFormat( o.getWorkbook() ) ).toBeTrue();
			});

			it( "Allows the workbook to be read post initialisation", function(){
				wb = s.newChainable().read( getTestFilePath( "test.xlsx" ) ).getWorkbook();
				expect( s.isXmlFormat( wb ) ).toBeTrue();
			});

			it( "Allows the workbook to be generated from a CSV file", function(){
				var csv = 'column1,column2#crlf#"Frumpo McNugget",12345';
				wb = s.newChainable().fromCsv( csv=csv, firstRowIsHeader=true ).getWorkbook();
				expect( s.getCellValue( wb, 2, 2 ) ).toBe( "12345" );
			});

			it( "Allows the workbook to be generated from a query", function(){
				var query = QueryNew( "Header1,Header2", "VarChar,VarChar",[ [ "a", "b" ],[ "c", "d" ] ] );
				wb = s.newChainable().fromQuery( query ).getWorkbook();
				actual = s.getSheetHelper().sheetToQuery( workbook=wb, headerRow=1 );
				expect( actual ).toBe( query );
			});

		});

		it( "Allows multiple operations on a single workbook object to be chained", function(){
			wb = s.newChainable( "xlsx" )
				.addRow( [ "a", "b", "c" ] )
				.formatCell( { bold=true }, 1, 1 )
				.getWorkbook();
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( "a" );
			expect( s.getCellFormat( wb, 1, 1 ).bold ).toBeTrue();
		});

		describe( "a chained call throws an exception if", function(){

			it( "no workbook has been passed in, read or initialised", function(){
				expect( function(){
					s.newChainable().addRow( [ "a", "b", "c" ] );
				}).toThrow( regex="Missing workbook" );
			});

			it( "the workbook is not a invalid object", function(){
				expect( function(){
					s.newChainable()
						.read( src=getTestFilePath( "test.xlsx" ), format="query" )
						.addRow( [ "a", "b", "c" ] );
				}).toThrow( regex="Invalid workbook" );
			});

		});

	});	

});	
</cfscript>