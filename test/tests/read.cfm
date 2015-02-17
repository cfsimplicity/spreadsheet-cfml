<cfscript>
describe( "read tests",function(){

	it( "can read a traditional XLS file",function() {
		path = ExpandPath( "/root/test/files/test.xls" );
		workbook = s.read( src=path );
		expect( workbook.getClass().name ).toBe( "org.apache.poi.hssf.usermodel.HSSFWorkbook" );
	});

	it( "can read an OOXML file",function() {
		path = ExpandPath( "/root/test/files/test.xlsx" );
		workbook = s.read( src=path );
		expect( workbook.getClass().name ).toBe( "org.apache.poi.xssf.usermodel.XSSFWorkbook" );
	});

	it( "can read a traditional XLS file into a query",function() {
		path = ExpandPath( "/root/test/files/test.xls" );
		expected = querySim(
			"column1,column2
			a|b
			c|d");
		actual = s.read( src=path,format="query" );
		expect( actual ).toBe( expected );
	});

	it( "can read an OOXML file into a query",function() {
		path = ExpandPath( "/root/test/files/test.xlsx" );
		expected = querySim(
			"column1,column2
			a|e
			b|f
			c|g
			I am|ooxml");
		actual = s.read( src=path,format="query" );
		expect( actual ).toBe( expected );
	});

	it( "reads from the specified sheet index",function(){
		path = ExpandPath( "/root/test/files/test.xls" );// has 2 sheets
		expected = querySim(
			"column1,column2
			x|y");
		actual = s.read( src=path,format="query",sheet=2 );
		expect( actual ).toBe( expected );
	});

	it( "uses header row for column names if specified",function() {
		path = ExpandPath( "/root/test/files/test.xls" );
		expected = querySim(
			"a,b
			c|d");
		actual = s.read( src=path,format="query",headerRow=1 );
		expect( actual ).toBe( expected );
	});

	it( "includes the specified header row in query if includeHeader is true",function() {
		path = ExpandPath( "/root/test/files/test.xls" );
		expected = querySim(
			"a,b
			a|b
			c|d");
		actual = s.read( src=path,format="query",headerRow=1,includeHeaderRow=true );
		expect( actual ).toBe( expected );
	});

	it( "includes blank rows in query if includeBlankRows is true",function() {
		data = QueryNew( "column1,column2","VarChar,VarChar",[ [ "","" ],[ "a","b" ] ] );
		workbook = s.new();
		s.addRows( workbook,data );
		s.write( workbook,tempXlsPath,true );
		expected = data;
		actual = s.read( src=tempXlsPath,format="query",includeBlankRows=true );
		expect( actual ).toBe( expected );
	});

});	
</cfscript>