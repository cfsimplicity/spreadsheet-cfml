<cfscript>
describe( "formatCell", function(){

	beforeEach( function(){
		xls = s.new();
		xlsx = s.newXlsx();
		s.setCellValue( xls, "test", 1, 1 );
		s.setCellValue( xlsx, "test", 1, 1 );
	});

	it( "can set bold", function(){
		var format = { bold: true };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.bold ).toBeTrue();
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.bold ).toBeTrue();
	});

	it( "can set the horizontal alignment", function(){
		var format = { alignment: "CENTER" };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.alignment ).toBe( "CENTER" );
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.alignment ).toBe( "CENTER" );
	});

	it( "can set the bottomborder", function(){
		var format = { bottomborder: "THICK" };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.bottomborder ).toBe( "THICK" );
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.bottomborder ).toBe( "THICK" );
	});

	it( "can set the bottombordercolor", function(){
		var format = { bottombordercolor: "RED" };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.bottombordercolor ).toBe( "255,0,0" );
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.bottombordercolor ).toBe( "255,0,0" );
	});

	it( "can set the color", function(){
		var format = { color: "BLUE" };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.color ).toBe( "0,0,255" );
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.color ).toBe( "0,0,255" );
	});

	it( "can set the dataformat", function(){
		var format = { dataformat: "@" };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.dataformat ).toBe( "@" );
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.dataformat ).toBe( "@" );
	});

	it( "can set the fgcolor", function(){
		var format = { fgcolor: "GREEN" };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.fgcolor ).toBe( "0,128,0" );
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.fgcolor ).toBe( "0,128,0" );
	});

	it( "can set the fillpattern", function(){
		var format = { fillpattern: "BRICKS" };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.fillpattern ).toBe( "BRICKS" );
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.fillpattern ).toBe( "BRICKS" );
	});

	it( "can set the font", function(){
		var format = { font: "Helvetica" };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.font ).toBe( "Helvetica" );
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.font ).toBe( "Helvetica" );
	});

	it( "can set the fontsize", function(){
		var format = { fontsize: 24 };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.fontsize ).toBe( 24 );
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.fontsize ).toBe( 24 );
	});

	it( "can set italic", function(){
		var format = { italic: true };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.italic ).toBeTrue();
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.italic ).toBeTrue();
	});

	it( "can set the leftborder", function(){
		var format = { leftborder: "THICK" };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.leftborder ).toBe( "THICK" );
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.leftborder ).toBe( "THICK" );
	});

	it( "can set the leftbordercolor", function(){
		var format = { leftbordercolor: "RED" };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.leftbordercolor ).toBe( "255,0,0" );
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.leftbordercolor ).toBe( "255,0,0" );
	});

	it( "can set quoteprefixed", function(){
		var formulaLikeString = "SUM(A2:A3)";
		var format = { quoteprefixed: true };
		var xls = s.new();
		s.addColumn( xls, "#formulaLikeString#,1,1" );
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.quoteprefixed ).toBeTrue();
		expect( s.getCellValue( xls, 1, 1 ) ).toBe( formulaLikeString );
		var xlsx = s.newXlsx();
		s.addColumn( xlsx, "#formulaLikeString#,1,1" );
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.quoteprefixed ).toBeTrue();
		expect( s.getCellValue( xlsx, 1, 1 ) ).toBe( formulaLikeString );
	});

	it( "can set the rightborder", function(){
		var format = { rightborder: "THICK" };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.rightborder ).toBe( "THICK" );
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.rightborder ).toBe( "THICK" );
	});

	it( "can set the rightbordercolor", function(){
		var format = { rightbordercolor: "RED" };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.rightbordercolor ).toBe( "255,0,0" );
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.rightbordercolor ).toBe( "255,0,0" );
	});

	it( "can set the rotation", function(){
		var format = { rotation: 90 };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.rotation ).toBe( 90 );
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.rotation ).toBe( 90 );
	});

	it( "can set strikeout", function(){
		var format = { strikeout: true };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.strikeout ).toBeTrue();
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.strikeout ).toBeTrue();
	});

	it( "can set textwrap", function(){
		var format = { textwrap: true };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.textwrap ).toBeTrue();
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.textwrap ).toBeTrue();
	});

	it( "can set the topborder", function(){
		var format = { topborder: "THICK" };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.topborder ).toBe( "THICK" );
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.topborder ).toBe( "THICK" );
	});

	it( "can set the topbordercolor", function(){
		var format = { topbordercolor: "RED" };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.topbordercolor ).toBe( "255,0,0" );
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.topbordercolor ).toBe( "255,0,0" );
	});

	it( "can set the vertical alignment", function(){
		var format = { verticalalignment: "CENTER" };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.verticalalignment ).toBe( "CENTER" );
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.verticalalignment ).toBe( "CENTER" );
	});

	it( "can set underline", function(){
		var format = { underline: true };
		s.formatCell( xls, format, 1, 1 );
		var cellFormat = s.getCellFormat( xls, 1, 1 );
		expect( cellFormat.underline ).toBeTrue();
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.underline ).toBeTrue();
	});

	it( "can set a non preset RGB triplet color on an XLSX workbook cell", function(){
		var triplet = "181,133,212";
		var format = { color: triplet, bottombordercolor: triplet };
		s.formatCell( xlsx, format, 1, 1 );
		var cellFormat = s.getCellFormat( xlsx, 1, 1 );
		expect( cellFormat.color ).toBe( triplet ); //font color
		expect( cellFormat.bottombordercolor ).toBe( triplet ); //style color
	});

});	
</cfscript>