<cfscript>
describe( "formatCell", ()=>{

	beforeEach( ()=>{
		s.clearCellStyleCache();
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( ( wb )=>{
			s.addColumn( wb, [ "a1", "a2" ] );
		})
	})

	setAndGetFormat = function( wb, format, overwriteCurrentStyle=true ){
		s.formatCell( workbook=wb, format=format, row=1, column=1, overwriteCurrentStyle=overwriteCurrentStyle );
		return s.getCellFormat( wb, 1, 1 );
	};

	it( "can get the format of an unformatted cell", ()=>{
		workbooks.Each( ( wb )=>{
			expect( s.getCellFormat( wb, 1, 1 ) ).toBeTypeOf( "struct" );
		})
	})

	it( "formatCell and getCellFormat are chainable", ()=>{
		var format = { bold: true };
		workbooks.Each( ( wb )=>{
			var cellFormat = s.newChainable( wb )
				.formatCell( format, 1, 1 )
				.getCellFormat( 1, 1 );
			expect( cellFormat.bold ).toBeTrue();
		})
	})

	it( "can set bold", ()=>{
		var format = { bold: true };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.bold ).toBeTrue();
		})
	})

	it( "can set the horizontal alignment", ()=>{
		var format = { alignment: "CENTER" };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.alignment ).toBe( "CENTER" );
		})
	})

	it( "can set the bottomborder", ()=>{
		var format = { bottomborder: "THICK" };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.bottomborder ).toBe( "THICK" );
		})
	})

	it( "can set the bottombordercolor", ()=>{
		var format = { bottombordercolor: "RED" };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.bottombordercolor ).toBe( "255,0,0" );
		})
	})

	it( "can set the color", ()=>{
		var format = { color: "BLUE" };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.color ).toBe( "0,0,255" );
		})
	})

	it( "can set the dataformat", ()=>{
		var format = { dataformat: "@" };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.dataformat ).toBe( "@" );
		})
	})

	it( "can set the fgcolor", ()=>{
		var format = { fgcolor: "GREEN" };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.fgcolor ).toBe( "0,128,0" );
		})
	})

	it( "will ensure a fillpattern is specified when setting the fgcolor", ()=>{
		var format = { fgcolor: "GREEN" };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.fgcolor ).toBe( "0,128,0" );
			expect( cellFormat.fillpattern ).toBe( "SOLID_FOREGROUND" );
		})
	})

	it( "can set the fillpattern", ()=>{
		var format = { fillpattern: "BRICKS" };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.fillpattern ).toBe( "BRICKS" );
		})
	})

	it( "can set the font", ()=>{
		var format = { font: "Helvetica" };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.font ).toBe( "Helvetica" );
		})
	})

	it( "can set the fontsize", ()=>{
		var format = { fontsize: 24 };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.fontsize ).toBe( 24 );
		})
	})

	it( "can set an indent of 15", ()=>{
		var format = { indent: 15 };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.indent ).toBe( 15 );
		})
	})

	it( "can set an indent of 15+ on XLSX", ()=>{
		var format = { indent: 17 };
		var xlsx = workbooks[ 2 ];
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.indent ).toBe( 17 );
	})

	it( "treats indents of 15+ as the maximum 15 on XLS", ()=>{
		var format = { indent: 17 };
		var xls = workbooks[ 1 ];
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.indent ).toBe( 15 );
	})

	it( "can set italic", ()=>{
		var format = { italic: true };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.italic ).toBeTrue();
		})
	})

	it( "can set the leftborder", ()=>{
		var format = { leftborder: "THICK" };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.leftborder ).toBe( "THICK" );
		})
	})

	it( "can set the leftbordercolor", ()=>{
		var format = { leftbordercolor: "RED" };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.leftbordercolor ).toBe( "255,0,0" );
		})
	})

	it( "can set quoteprefixed", ()=>{
		var formulaLikeString = "SUM(A2:A3)";
		var format = { quoteprefixed: true };
		workbooks.Each( ( wb )=>{
			s.setCellValue( wb, formulaLikeString, 1, 1 );
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.quoteprefixed ).toBeTrue();
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( formulaLikeString );
		})
	})

	it( "can set the rightborder", ()=>{
		var format = { rightborder: "THICK" };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.rightborder ).toBe( "THICK" );
		})
	})

	it( "can set the rightbordercolor", ()=>{
		var format = { rightbordercolor: "RED" };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.rightbordercolor ).toBe( "255,0,0" );
		})
	})

	it( "can set the rotation", ()=>{
		var format = { rotation: 90 };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.rotation ).toBe( 90 );
		})
	})

	it( "can set strikeout", ()=>{
		var format = { strikeout: true };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.strikeout ).toBeTrue();
		})
	})

	it( "can set textwrap", ()=>{
		var format = { textwrap: true };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.textwrap ).toBeTrue();
		})
	})

	it( "can set the topborder", ()=>{
		var format = { topborder: "THICK" };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.topborder ).toBe( "THICK" );
		})
	})

	it( "can set the topbordercolor", ()=>{
		var format = { topbordercolor: "RED" };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.topbordercolor ).toBe( "255,0,0" );
		})
	})

	it( "can set the vertical alignment", ()=>{
		var format = { verticalalignment: "CENTER" };
		workbooks.Each( ( wb )=>{
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.verticalalignment ).toBe( "CENTER" );
		})
	})

	it( "can set underline as boolean", ()=>{
		workbooks.Each( ( wb )=>{
			s.formatCell( wb, { underline: true }, 1, 1 );
			var cellFormat = s.getCellFormat( wb, 1, 1 );
			expect( cellFormat.underline ).toBe( "single" );
			//check turning off
			s.formatCell( wb, { underline: false }, 1, 1 );
			var cellFormat = s.getCellFormat( wb, 1, 1 );
			expect( cellFormat.underline ).toBe( "none" );
		})
	})

	it( "can set underline as 'single' or 'none'", ()=>{
		workbooks.Each( ( wb )=>{
			s.formatCell( wb, { underline: "single" }, 1, 1 );
			var cellFormat = s.getCellFormat( wb, 1, 1 );
			expect( cellFormat.underline ).toBe( "single" );
			//check turning off
			s.formatCell( wb, { underline: "none" }, 1, 1 );
			var cellFormat = s.getCellFormat( wb, 1, 1 );
			expect( cellFormat.underline ).toBe( "none" );
		})
	})

	it( "can set underline as 'double'", ()=>{
		workbooks.Each( ( wb )=>{
			s.formatCell( wb, { underline: "double" }, 1, 1 );
			var cellFormat = s.getCellFormat( wb, 1, 1 );
			expect( cellFormat.underline ).toBe( "double" );
		})
	})

	it( "can set underline as 'single accounting'", ()=>{
		workbooks.Each( ( wb )=>{
			s.formatCell( wb, { underline: "single accounting" }, 1, 1 );
			var cellFormat = s.getCellFormat( wb, 1, 1 );
			expect( cellFormat.underline ).toBe( "single accounting" );
		})
	})

	it( "can set underline as 'double accounting'", ()=>{
		workbooks.Each( ( wb )=>{
			s.formatCell( wb, { underline: "double accounting" }, 1, 1 );
			var cellFormat = s.getCellFormat( wb, 1, 1 );
			expect( cellFormat.underline ).toBe( "double accounting" );
		})
	})

	it( "ignores an invalid underline value", ()=>{
		workbooks.Each( ( wb )=>{
			s.formatCell( wb, { underline: "blah" }, 1, 1 );
			var cellFormat = s.getCellFormat( wb, 1, 1 );
			expect( cellFormat.underline ).toBe( "none" );
		})
	})

	it( "will map 9 deprecated colour names ending in 1 to the corresponding valid value", ()=>{
		workbooks.Each( ( wb )=>{
			// include numbers either side of 127 which might throw ACF
			var deprecatedName = "RED1";
			var format = { color: deprecatedName, bottombordercolor: deprecatedName };
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.color ).toBe( "255,0,0" ); //font color
			expect( cellFormat.bottombordercolor ).toBe( "255,0,0" ); //style color
		})
	})

	it( "can set a non preset RGB triplet color on an XLSX workbook cell", ()=>{
		var xlsx = workbooks[ 2 ];
		// include numbers either side of 127 which might throw ACF
		var triplet = "64,255,128";
		var format = { color: triplet, bottombordercolor: triplet };
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.color ).toBe( triplet ); //font color
		expect( cellFormat.bottombordercolor ).toBe( triplet ); //style color
	})

	it( "can set a non preset Hex color value on an XLSX workbook cell", ()=>{
		var xlsx = workbooks[ 2 ];
		var hex = "FFFFFF";
		var triplet = "255,255,255";
		var format = { color: hex, bottombordercolor: hex };
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.color ).toBe( triplet ); //font color
		expect( cellFormat.bottombordercolor ).toBe( triplet ); //style color
		// handle leading #
		var hex = "##FFFFFF";
		var triplet = "255,255,255";
		var format = { color: hex, bottombordercolor: hex };
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.color ).toBe( triplet ); //font color
		expect( cellFormat.bottombordercolor ).toBe( triplet ); //style color
	})

	it( "can preserve the existing font properties and not affect other cells", ()=>{
		workbooks.Each( ( wb )=>{
			var cellA2originalFormat = s.getCellFormat( wb, 2, 1 );
			var format = { font: "Helvetica" };
			setAndGetFormat( wb, format, false );
			expect( s.getCellFormat( wb, 2, 1 ).font ).toBe( cellA2originalFormat.font );//should be unchanged
			format = { bold: true };
			var cellFormat = setAndGetFormat( wb, format, false );
			expect( cellFormat.font ).toBe( "Helvetica" );
			format = { color: "BLUE" };
			cellFormat = setAndGetFormat( wb, format, false );
			expect( cellFormat.font ).toBe( "Helvetica" );
			format = { fontsize: 24 };
			cellFormat = setAndGetFormat( wb, format, false );
			expect( cellFormat.font ).toBe( "Helvetica" );
			format = { italic: true };
			cellFormat = setAndGetFormat( wb, format, false );
			expect( cellFormat.font ).toBe( "Helvetica" );
			format = { strikeout: true };
			cellFormat = setAndGetFormat( wb, format, false );
			expect( cellFormat.font ).toBe( "Helvetica" );
			format = { underline: true };
			cellFormat = setAndGetFormat( wb, format, false );
			expect( cellFormat.font ).toBe( "Helvetica" );
			format = { font: "Courier New" };
			cellFormat = setAndGetFormat( wb, format, false );
			expect( cellFormat.font ).toBe( "Courier New" );
			expect( cellFormat.fontsize ).toBe( 24 );
		})
	})

	it( "allows styles to be set using a pre-built cellStyle object", ()=>{
		workbooks.Each( ( wb )=>{
			expect( s.getCellFormat( wb, 1, 1 ).bold ).toBeFalse();
			var cellStyle = s.createCellStyle( wb, { bold: true } );
			// format argument may be a cellStyle object instead of a struct
			s.formatCell( wb, cellStyle, 1, 1 );
			expect( s.getCellFormat( wb, 1, 1 ).bold ).toBeTrue();
			// support previous separate cellStyle argument without format
			expect( s.getCellFormat( wb, 2, 1 ).bold ).toBeFalse();
			s.formatCell( workbook=wb, row=2, column=1, cellStyle=cellStyle );
			expect( s.getCellFormat( wb, 2, 1 ).bold ).toBeTrue();
		})
	})

	it( "caches and re-uses indentical formats passed as a struct over the life of the library to avoid cell style duplication", ()=>{
		variables.data = [ [ "x", "y" ] ];
		variables.format = { bold: true };
		workbooks.Each( ( wb )=>{
			s.addRows( wb, data );
			var expected = s.isXmlFormat( wb )? 1: 21;
			expect( s.getWorkbookCellStylesTotal( wb ) ).toBe( expected );
			cfloop( from=1, to=2, index="local.row" ){
				s.formatCell( wb, format, row, 1 );
			}
			expected = s.isXmlFormat( wb )? 2: 22;
			expect( s.getWorkbookCellStylesTotal( wb ) ).toBe( expected );
		})
	})

	describe( "formatCell throws an exception if", ()=>{

		it( "neither a format struct nor a cellStyle are passed", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					s.formatCell( workbook=wb, row=1, column=1 );
				}).toThrow( type="cfsimplicity.spreadsheet.missingRequiredArgument" );
			})
		})

		it( "an invalid hex value is passed", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					var hex = "GGHHII";
					var format = { color: hex, bottombordercolor: hex };
					var cellFormat = setAndGetFormat( wb, format );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidColor" );
			})
		})

		it( "a cellStyle object is passed with the overwriteCurrentStyle flag set to false", ()=>{
			workbooks.Each( ( wb )=>{
				expect( ()=>{
					var cellStyle = s.createCellStyle( wb, { bold: true } );
					s.formatCell( workbook=wb, format=cellStyle, row=1, column=1, overwriteCurrentStyle=false );
				}).toThrow( type="cfsimplicity.spreadsheet.invalidArgumentCombination" );
			})
		})

	})

})	
</cfscript>