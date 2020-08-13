<cfscript>
describe( "formatCell", function(){

	beforeEach( function(){
		xls = s.new();
		xlsx = s.newXlsx();
		s.setCellValue( xls, "test", 1, 1 );
		s.setCellValue( xlsx, "test", 1, 1 );
	});

	setAndGetFormat = function( wb, format, overwriteCurrentStyle=true ){
		s.formatCell( workbook=wb, format=format, row=1, column=1, overwriteCurrentStyle=overwriteCurrentStyle );
		return s.getCellFormat( wb, 1, 1 );
	};

	it( "can set bold", function(){
		var format = { bold: true };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.bold ).toBeTrue();
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.bold ).toBeTrue();
	});

	it( "can set the horizontal alignment", function(){
		var format = { alignment: "CENTER" };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.alignment ).toBe( "CENTER" );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.alignment ).toBe( "CENTER" );
	});

	it( "can set the bottomborder", function(){
		var format = { bottomborder: "THICK" };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.bottomborder ).toBe( "THICK" );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.bottomborder ).toBe( "THICK" );
	});

	it( "can set the bottombordercolor", function(){
		var format = { bottombordercolor: "RED" };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.bottombordercolor ).toBe( "255,0,0" );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.bottombordercolor ).toBe( "255,0,0" );
	});

	it( "can set the color", function(){
		var format = { color: "BLUE" };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.color ).toBe( "0,0,255" );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.color ).toBe( "0,0,255" );
	});

	it( "can set the dataformat", function(){
		var format = { dataformat: "@" };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.dataformat ).toBe( "@" );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.dataformat ).toBe( "@" );
	});

	it( "can set the fgcolor", function(){
		var format = { fgcolor: "GREEN" };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.fgcolor ).toBe( "0,128,0" );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.fgcolor ).toBe( "0,128,0" );
	});

	it( "will ensure a fillpattern is specified when setting the fgcolor", function(){
		var format = { fgcolor: "GREEN" };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.fgcolor ).toBe( "0,128,0" );
		expect( cellFormat.fillpattern ).toBe( "SOLID_FOREGROUND" );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.fgcolor ).toBe( "0,128,0" );
		expect( cellFormat.fillpattern ).toBe( "SOLID_FOREGROUND" );
	});

	it( "can set the fillpattern", function(){
		var format = { fillpattern: "BRICKS" };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.fillpattern ).toBe( "BRICKS" );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.fillpattern ).toBe( "BRICKS" );
	});

	it( "can set the font", function(){
		var format = { font: "Helvetica" };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.font ).toBe( "Helvetica" );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.font ).toBe( "Helvetica" );
	});

	it( "can set the fontsize", function(){
		var format = { fontsize: 24 };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.fontsize ).toBe( 24 );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.fontsize ).toBe( 24 );
	});

	it( "can set an indent of 15", function(){
		var format = { indent: 15 };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.indent ).toBe( 15 );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.indent ).toBe( 15 );
	});

	it( "can set an indent of 15+ on XLSX", function(){
		var format = { indent: 17 };
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.indent ).toBe( 17 );
	});

	it( "treats indents of 15+ as the maximum 15 on XLS", function(){
		var format = { indent: 17 };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.indent ).toBe( 15 );
	});

	it( "can set italic", function(){
		var format = { italic: true };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.italic ).toBeTrue();
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.italic ).toBeTrue();
	});

	it( "can set the leftborder", function(){
		var format = { leftborder: "THICK" };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.leftborder ).toBe( "THICK" );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.leftborder ).toBe( "THICK" );
	});

	it( "can set the leftbordercolor", function(){
		var format = { leftbordercolor: "RED" };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.leftbordercolor ).toBe( "255,0,0" );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.leftbordercolor ).toBe( "255,0,0" );
	});

	it( "can set quoteprefixed", function(){
		var formulaLikeString = "SUM(A2:A3)";
		var format = { quoteprefixed: true };
		var xls = s.new();
		s.addColumn( xls, "#formulaLikeString#,1,1" );
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.quoteprefixed ).toBeTrue();
		expect( s.getCellValue( xls, 1, 1 ) ).toBe( formulaLikeString );
		var xlsx = s.newXlsx();
		s.addColumn( xlsx, "#formulaLikeString#,1,1" );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.quoteprefixed ).toBeTrue();
		expect( s.getCellValue( xlsx, 1, 1 ) ).toBe( formulaLikeString );
	});

	it( "can set the rightborder", function(){
		var format = { rightborder: "THICK" };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.rightborder ).toBe( "THICK" );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.rightborder ).toBe( "THICK" );
	});

	it( "can set the rightbordercolor", function(){
		var format = { rightbordercolor: "RED" };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.rightbordercolor ).toBe( "255,0,0" );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.rightbordercolor ).toBe( "255,0,0" );
	});

	it( "can set the rotation", function(){
		var format = { rotation: 90 };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.rotation ).toBe( 90 );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.rotation ).toBe( 90 );
	});

	it( "can set strikeout", function(){
		var format = { strikeout: true };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.strikeout ).toBeTrue();
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.strikeout ).toBeTrue();
	});

	it( "can set textwrap", function(){
		var format = { textwrap: true };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.textwrap ).toBeTrue();
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.textwrap ).toBeTrue();
	});

	it( "can set the topborder", function(){
		var format = { topborder: "THICK" };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.topborder ).toBe( "THICK" );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.topborder ).toBe( "THICK" );
	});

	it( "can set the topbordercolor", function(){
		var format = { topbordercolor: "RED" };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.topbordercolor ).toBe( "255,0,0" );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.topbordercolor ).toBe( "255,0,0" );
	});

	it( "can set the vertical alignment", function(){
		var format = { verticalalignment: "CENTER" };
		var cellFormat = setAndGetFormat( xls, format );
		expect( cellFormat.verticalalignment ).toBe( "CENTER" );
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.verticalalignment ).toBe( "CENTER" );
	});

	it( "can set underline as boolean", function(){
		s.formatCell( xls, { underline: true }, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.underline ).toBe( "single" );
		s.formatCell( xlsx, { underline: true }, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.underline ).toBe( "single" );
		//check turning off
		s.formatCell( xls, { underline: false }, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.underline ).toBe( "none" );
		s.formatCell( xlsx, { underline: false }, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.underline ).toBe( "none" );
	});

	it( "can set underline as 'single' or 'none'", function(){
		s.formatCell( xls, { underline: "single" }, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.underline ).toBe( "single" );
		s.formatCell( xlsx, { underline: "single" }, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.underline ).toBe( "single" );
		//check turning off
		s.formatCell( xls, { underline: "none" }, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.underline ).toBe( "none" );
		s.formatCell( xlsx, { underline: "none" }, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.underline ).toBe( "none" );
	});

	it( "can set underline as 'double'", function(){
		s.formatCell( xls, { underline: "double" }, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.underline ).toBe( "double" );
		s.formatCell( xlsx, { underline: "double" }, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.underline ).toBe( "double" );
	});

	it( "can set underline as 'single accounting'", function(){
		s.formatCell( xls, { underline: "single accounting" }, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.underline ).toBe( "single accounting" );
		s.formatCell( xlsx, { underline: "single accounting" }, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.underline ).toBe( "single accounting" );
	});

	it( "can set underline as 'double accounting'", function(){
		s.formatCell( xls, { underline: "double accounting" }, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.underline ).toBe( "double accounting" );
		s.formatCell( xlsx, { underline: "double accounting" }, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.underline ).toBe( "double accounting" );
	});

	it( "will map 9 deprecated colour names ending in 1 to the corresponding valid value", function(){
		// include numbers either side of 127 which might throw ACF
		var deprecatedName = "RED1";
		var format = { color: deprecatedName, bottombordercolor: deprecatedName };
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.color ).toBe( "255,0,0" ); //font color
		expect( cellFormat.bottombordercolor ).toBe( "255,0,0" ); //style color
	});

	it( "can set a non preset RGB triplet color on an XLSX workbook cell", function(){
		// include numbers either side of 127 which might throw ACF
		var triplet = "64,255,128";
		var format = { color: triplet, bottombordercolor: triplet };
		var cellFormat = setAndGetFormat( xlsx, format );
		expect( cellFormat.color ).toBe( triplet ); //font color
		expect( cellFormat.bottombordercolor ).toBe( triplet ); //style color
	});

	it( "can set a non preset Hex color value on an XLSX workbook cell", function(){
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


	it( "can preserve the existing font properties when setting bold, color, font name, font size, italic, strikeout and underline", function(){
		var format = { font: "Helvetica" };
		setAndGetFormat( xls, format );
		format = { bold: true };
		var cellFormat = setAndGetFormat( xls, format, false );
		expect( cellFormat.font ).toBe( "Helvetica" );
		var format = { color: "BLUE" };
		var cellFormat = setAndGetFormat( xls, format, false );
		expect( cellFormat.font ).toBe( "Helvetica" );
		var format = { fontsize: 24 };
		var cellFormat = setAndGetFormat( xls, format, false );
		expect( cellFormat.font ).toBe( "Helvetica" );
		var format = { italic: true };
		var cellFormat = setAndGetFormat( xls, format, false );
		expect( cellFormat.font ).toBe( "Helvetica" );
		var format = { strikeout: true };
		var cellFormat = setAndGetFormat( xls, format, false );
		expect( cellFormat.font ).toBe( "Helvetica" );
		var format = { underline: true };
		var cellFormat = setAndGetFormat( xls, format, false );
		expect( cellFormat.font ).toBe( "Helvetica" );
		var format = { font: "Courier New" };
		var cellFormat = setAndGetFormat( xls, format, false );
		expect( cellFormat.fontsize ).toBe( 24 );
	});

});	
</cfscript>