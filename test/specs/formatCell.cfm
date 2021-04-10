<cfscript>
describe( "formatCell", function(){

	beforeEach( function(){
		variables.workbooks = [ s.newXls(), s.newXlsx() ];
		workbooks.Each( function( wb ){
			s.setCellValue( wb, "test", 1, 1 );
		});
	});

	setAndGetFormat = function( wb, format, overwriteCurrentStyle=true ){
		s.formatCell( workbook=wb, format=format, row=1, column=1, overwriteCurrentStyle=overwriteCurrentStyle );
		return s.getCellFormat( wb, 1, 1 );
	};

	it( "can set bold", function(){
		var format = { bold: true };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.bold ).toBeTrue();
		});
	});

	it( "can set the horizontal alignment", function(){
		var format = { alignment: "CENTER" };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.alignment ).toBe( "CENTER" );
		});
	});

	it( "can set the bottomborder", function(){
		var format = { bottomborder: "THICK" };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.bottomborder ).toBe( "THICK" );
		});
	});

	it( "can set the bottombordercolor", function(){
		var format = { bottombordercolor: "RED" };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.bottombordercolor ).toBe( "255,0,0" );
		});
	});

	it( "can set the color", function(){
		var format = { color: "BLUE" };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.color ).toBe( "0,0,255" );
		});
	});

	it( "can set the dataformat", function(){
		var format = { dataformat: "@" };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.dataformat ).toBe( "@" );
		});
	});

	it( "can set the fgcolor", function(){
		var format = { fgcolor: "GREEN" };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.fgcolor ).toBe( "0,128,0" );
		});
	});

	it( "will ensure a fillpattern is specified when setting the fgcolor", function(){
		var format = { fgcolor: "GREEN" };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.fgcolor ).toBe( "0,128,0" );
			expect( cellFormat.fillpattern ).toBe( "SOLID_FOREGROUND" );
		});
	});

	it( "can set the fillpattern", function(){
		var format = { fillpattern: "BRICKS" };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.fillpattern ).toBe( "BRICKS" );
		});
	});

	it( "can set the font", function(){
		var format = { font: "Helvetica" };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.font ).toBe( "Helvetica" );
		});
	});

	it( "can set the fontsize", function(){
		var format = { fontsize: 24 };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.fontsize ).toBe( 24 );
		});
	});

	it( "can set an indent of 15", function(){
		var format = { indent: 15 };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.indent ).toBe( 15 );
		});
	});

	it( "can set an indent of 15+ on XLSX", function(){
		var format = { indent: 17 };
		var xlsx = workbooks[ 2 ];
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.indent ).toBe( 17 );
	});

	it( "treats indents of 15+ as the maximum 15 on XLS", function(){
		var format = { indent: 17 };
		var xls = workbooks[ 1 ];
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.indent ).toBe( 15 );
	});

	it( "can set italic", function(){
		var format = { italic: true };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.italic ).toBeTrue();
		});
	});

	it( "can set the leftborder", function(){
		var format = { leftborder: "THICK" };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.leftborder ).toBe( "THICK" );
		});
	});

	it( "can set the leftbordercolor", function(){
		var format = { leftbordercolor: "RED" };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.leftbordercolor ).toBe( "255,0,0" );
		});
	});

	it( "can set quoteprefixed", function(){
		var formulaLikeString = "SUM(A2:A3)";
		var format = { quoteprefixed: true };
		workbooks.Each( function( wb ){
			s.addColumn( wb, formulaLikeString, 1, 1 );
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.quoteprefixed ).toBeTrue();
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( formulaLikeString );
		});
	});

	it( "can set the rightborder", function(){
		var format = { rightborder: "THICK" };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.rightborder ).toBe( "THICK" );
		});
	});

	it( "can set the rightbordercolor", function(){
		var format = { rightbordercolor: "RED" };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.rightbordercolor ).toBe( "255,0,0" );
		});
	});

	it( "can set the rotation", function(){
		var format = { rotation: 90 };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.rotation ).toBe( 90 );
		});
	});

	it( "can set strikeout", function(){
		var format = { strikeout: true };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.strikeout ).toBeTrue();
		});
	});

	it( "can set textwrap", function(){
		var format = { textwrap: true };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.textwrap ).toBeTrue();
		});
	});

	it( "can set the topborder", function(){
		var format = { topborder: "THICK" };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.topborder ).toBe( "THICK" );
		});
	});

	it( "can set the topbordercolor", function(){
		var format = { topbordercolor: "RED" };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.topbordercolor ).toBe( "255,0,0" );
		});
	});

	it( "can set the vertical alignment", function(){
		var format = { verticalalignment: "CENTER" };
		workbooks.Each( function( wb ){
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.verticalalignment ).toBe( "CENTER" );
		});
	});

	it( "can set underline as boolean", function(){
		workbooks.Each( function( wb ){
			s.formatCell( wb, { underline: true }, 1, 1 );
			var cellFormat = s.getCellFormat( wb, 1, 1 );
			expect( cellFormat.underline ).toBe( "single" );
			//check turning off
			s.formatCell( wb, { underline: false }, 1, 1 );
			var cellFormat = s.getCellFormat( wb, 1, 1 );
			expect( cellFormat.underline ).toBe( "none" );
		});
	});

	it( "can set underline as 'single' or 'none'", function(){
		workbooks.Each( function( wb ){
			s.formatCell( wb, { underline: "single" }, 1, 1 );
			var cellFormat = s.getCellFormat( wb, 1, 1 );
			expect( cellFormat.underline ).toBe( "single" );
			//check turning off
			s.formatCell( wb, { underline: "none" }, 1, 1 );
			var cellFormat = s.getCellFormat( wb, 1, 1 );
			expect( cellFormat.underline ).toBe( "none" );
		});
	});

	it( "can set underline as 'double'", function(){
		workbooks.Each( function( wb ){
			s.formatCell( wb, { underline: "double" }, 1, 1 );
			var cellFormat = s.getCellFormat( wb, 1, 1 );
			expect( cellFormat.underline ).toBe( "double" );
		});
	});

	it( "can set underline as 'single accounting'", function(){
		workbooks.Each( function( wb ){
			s.formatCell( wb, { underline: "single accounting" }, 1, 1 );
			var cellFormat = s.getCellFormat( wb, 1, 1 );
			expect( cellFormat.underline ).toBe( "single accounting" );
		});
	});

	it( "can set underline as 'double accounting'", function(){
		workbooks.Each( function( wb ){
			s.formatCell( wb, { underline: "double accounting" }, 1, 1 );
			var cellFormat = s.getCellFormat( wb, 1, 1 );
			expect( cellFormat.underline ).toBe( "double accounting" );
		});
	});

	it( "will map 9 deprecated colour names ending in 1 to the corresponding valid value", function(){
		workbooks.Each( function( wb ){
			// include numbers either side of 127 which might throw ACF
			var deprecatedName = "RED1";
			var format = { color: deprecatedName, bottombordercolor: deprecatedName };
			var cellFormat = setAndGetFormat( wb, format );
			expect( cellFormat.color ).toBe( "255,0,0" ); //font color
			expect( cellFormat.bottombordercolor ).toBe( "255,0,0" ); //style color
		});
	});

	it( "can set a non preset RGB triplet color on an XLSX workbook cell", function(){
		var xlsx = workbooks[ 2 ];
		// include numbers either side of 127 which might throw ACF
		var triplet = "64,255,128";
		var format = { color: triplet, bottombordercolor: triplet };
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.color ).toBe( triplet ); //font color
		expect( cellFormat.bottombordercolor ).toBe( triplet ); //style color
	});

	it( "can set a non preset Hex color value on an XLSX workbook cell", function(){
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
	});

	it( "Throws an exception if an invalid hex value is passed", function(){
		workbooks.Each( function( wb ){
			expect( function(){
				var hex = "GGHHII";
				var format = { color: hex, bottombordercolor: hex };
				var cellFormat = setAndGetFormat( wb, format );
			}).toThrow( regex="Invalid color" );
		});
	});


	it( "can preserve the existing font properties when setting bold, color, font name, font size, italic, strikeout and underline", function(){
		workbooks.Each( function( wb ){
			var format = { font: "Helvetica" };
			setAndGetFormat( wb, format );
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
		});
	});

});	
</cfscript>