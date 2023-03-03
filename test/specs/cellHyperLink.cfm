<cfscript>
describe( "cellHyperLinks", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		variables.uri = "https://w3c.org";
	});

	it( "setCellHyperLink and getCellHyperLink are chainable", function(){
		workbooks.Each( function( wb ){
			var actual = s.newChainable( wb )
				.setCellHyperlink( uri, 1, 1 )
				.getCellHyperlink( 1, 1 );
			expect( actual ).toBe( uri );
		});
	});

	describe( "getCellHyperlink", function(){

		it( "returns the address/URL of a cell's hyperlink", function(){
			workbooks.Each( function( wb ){
				s.setCellHyperlink( wb, uri, 1, 1 );
				expect( s.getCellHyperlink( wb, 1, 1 ) ).toBe( uri );
			});
		});

		it( "returns an empty string if the cell contains no hyperlink", function(){
			workbooks.Each( function( wb ){
				expect( s.getCellHyperlink( wb, 1, 1 ) ).toBeEmpty();
			});
		});

	});

	describe( "setHyperlink", function(){

		it( "adds a hyperlink to a cell", function(){
			workbooks.Each( function( wb ){
				s.setCellHyperlink( wb, uri, 1, 1 );
				expect( s.getCellHyperlink( wb, 1, 1 ) ).toBe( uri );
			});
		});

		it( "Allows the cell value to be specified", function(){
			workbooks.Each( function( wb ){
				var value = "W3C";
				s.setCellHyperlink( workbook=wb, row=1, column=1, link=uri, cellValue=value );
				expect( s.getCellHyperlink( wb, 1, 1 ) ).toBe( uri );
				expect( s.getCellValue( wb, 1, 1 ) ).toBe( value );
			});
		});

		it( "formats the hyperlink as blue/underlined by default", function(){
			workbooks.Each( function( wb ){
				s.setCellHyperlink( wb, uri, 1, 1 );
				var cellFormat = s.getCellFormat( wb, 1, 1 );
				expect( cellFormat.underline ).toBe( "single" );
				expect( cellFormat.color ).toBe( "0,0,255" );
			});
		});

		it( "by default re-uses the default hyperlink cell style across multiple calls", function(){
			workbooks.Each( function( wb ){
				//baseline
				var expectedNumberOfStyles = s.isXmlFormat( wb )? 1: 21;
				expect( s.getWorkbookCellStylesTotal( wb ) ).toBe( expectedNumberOfStyles );
				// first call
				s.setCellHyperlink( wb, uri, 1, 1 );
				expectedNumberOfStyles = s.isXmlFormat( wb )? 2: 22;
				expect( s.getWorkbookCellStylesTotal( wb ) ).toBe( expectedNumberOfStyles );
				// second call
				s.setCellHyperlink( wb, uri, 1, 2 );
				expect( s.getWorkbookCellStylesTotal( wb ) ).toBe( expectedNumberOfStyles );
			});
		});

		it( "allows hyperlink formatting to be overridden", function(){
			workbooks.Each( function( wb ){
				var format = { color: "RED", underline: false };
				s.setCellHyperlink( workbook=wb, row=1, column=1, link=uri, format=format );
				var cellFormat = s.getCellFormat( wb, 1, 1 );
				expect( cellFormat.underline ).toBe( "none" );
				expect( cellFormat.color ).toBe( "255,0,0" );
				//no formatting
				s.setCellHyperlink( workbook=wb, row=1, column=2, link=uri, format={} );
				var cellFormat = s.getCellFormat( wb, 1, 2 );
				expect( cellFormat.underline ).toBe( "none" );
			});
		});

		it( "allows hyperlink formatting to be overridden using a re-usable cellStyle", function(){
			workbooks.Each( function( wb ){
				var format = { color: "RED", underline: false };
				var cellStyle = s.createCellStyle( wb, format );
				s.setCellHyperlink( workbook=wb, row=1, column=1, link=uri, format=cellStyle );
				var cellFormat = s.getCellFormat( wb, 1, 1 );
				expect( cellFormat.underline ).toBe( "none" );
				expect( cellFormat.color ).toBe( "255,0,0" );
			});
		});

		it( "Allows email links to be added", function(){
			workbooks.Each( function( wb ){
				var email = "mailto:test@example.com";
				s.setCellHyperlink( workbook=wb, row=1, column=1, link=email, type="email" );
				expect( s.getCellHyperlink( wb, 1, 1 ) ).toBe( email );
				expect( s.getCellHelper().getCellAt( wb, 1, 1 ).getHyperLink().getType().name() ).toBe( "EMAIL" );
			});
		});

		it( "Allows file links to be added", function(){
			workbooks.Each( function( wb ){
				var file = "linked.xlsx";
				s.setCellHyperlink( workbook=wb, row=1, column=1, link=file, type="file" );
				expect( s.getCellHyperlink( wb, 1, 1 ) ).toBe( file );
				expect( s.getCellHelper().getCellAt( wb, 1, 1 ).getHyperLink().getType().name() ).toBe( "FILE" );
			});
		});

		it( "Allows internal links to be added", function(){
			workbooks.Each( function( wb ){
				var link = "'Target Sheet'!A1";
				s.setCellHyperlink( workbook=wb, row=1, column=1, link=link, type="document" );
				expect( s.getCellHyperlink( wb, 1, 1 ) ).toBe( link );
				expect( s.getCellHelper().getCellAt( wb, 1, 1 ).getHyperLink().getType().name() ).toBe( "DOCUMENT" );
			});
		});

		it( "Allows xlsx sheet hyperlink tooltips to be set", function(){
			var wb = s.newXlsx();
			var tooltip = "I'm a tooltip";
			s.setCellHyperlink( workbook=wb, row=1, column=1, link=uri, tooltip=tooltip );
			expect( s.getCellHyperlink( wb, 1, 1 ) ).toBe( uri );
			expect( s.getCellHelper().getCellAt( wb, 1, 1 ).getHyperLink().getTooltip() ).toBe( tooltip );
		});

		describe( "setCellHyperlink throws an exception if", function(){

			it( "an invalid type value is specified", function(){
				expect( function(){
					s.setCellHyperlink( workbook=variables.workbooks[ 1 ], row=1, column=1, link="https://w3c.org", type="blah" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidTypeArgument" );
			});

			it( "the workbook is XLS and a tooltip is specified", function(){
				expect( function(){
					s.setCellHyperlink( workbook=variables.workbooks[ 1 ], row=1, column=1, link=uri, tooltip="whatever" );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidSpreadsheetType" );
			});

		});

	});

});
</cfscript>