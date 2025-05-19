<cfscript>
describe( "read throws an exception if", ()=>{

	it( "queryColumnTypes is specified as a 'columnName/type' struct, but headerRow and columnNames arguments are missing", ()=>{
		expect( ()=>{
			var columnTypes = { col1: "Integer" };
			s.read( src=getTestFilePath( "test.xlsx" ), format="query", queryColumnTypes=columnTypes );
		}).toThrow( type="cfsimplicity.spreadsheet.invalidQueryColumnTypesArgument" );
	})

	it( "the 'query' argument is passed", ()=>{
		expect( ()=>{
			s.read( src=tempXlsPath, query="q" );
		}).toThrow( type="cfsimplicity.spreadsheet.invalidQueryArgument" );
	})

	it( "the format argument is invalid", ()=>{
		expect( ()=>{
			s.read( src=tempXlsPath, format="wrong" );
		}).toThrow( type="cfsimplicity.spreadsheet.invalidReadFormat" );
	})

	it( "the file doesn't exist", ()=>{
		expect( ()=>{
			var path = getTestFilePath( "nonexistent.xls" );
			s.read( src=path );
		}).toThrow( type="cfsimplicity.spreadsheet.nonExistentFile" );
	})

	it( "the sheet name doesn't exist", ()=>{
		expect( ()=>{
			var path = getTestFilePath( "test.xls" );
			s.read( src=path, format="query", sheetName="nonexistent" );
		}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetName" );
	})

	it( "the sheet number doesn't exist", ()=>{
		expect( ()=>{
			var path = getTestFilePath( "test.xls" );
			s.read( src=path, format="query", sheetNumber=20 );
		}).toThrow( type="cfsimplicity.spreadsheet.invalidSheetNumber" );
	})

	it( "both sheetName and sheetNumber arguments are specified", ()=>{
		expect( ()=>{
			var path = getTestFilePath( "test.xls" );
			s.read( src=path, sheetName="sheet1", sheetNumber=2 );
		}).toThrow( type="cfsimplicity.spreadsheet.invalidArguments" );
	})

	it( "the password for an encrypted XML file is incorrect", ()=>{
		expect( ()=>{
			var tempXlsxPath = getTestFilePath( "passworded.xlsx" );
			s.read( src=tempXlsxPath, format="query", password="parse" );
		}).toThrow( regex="(Invalid password|Password incorrect|password is invalid)" );
	})

	it( "the password for an encrypted binary file is incorrect", ()=>{
		expect( ()=>{
			var xlsPath = getTestFilePath( "passworded.xls" );
			s.read( src=xlsPath, format="query", password="parse" );
		}).toThrow( regex="(Invalid password|Password incorrect|password is invalid)" );
	})

	it( "the source file is not a spreadsheet", ()=>{
		expect( ()=>{
			var path = getTestFilePath( "notaspreadsheet.txt" );
			s.read( src=path );
		}).toThrow( type="cfsimplicity.spreadsheet.invalidFile" );
	})

	it( "the source file appears to contain CSV or TSV, and suggests using 'csvToQuery'", ()=>{
		expect( ()=>{
			var path = getTestFilePath( "csv.xls" );
			s.read( src=path );
		}).toThrow( type="cfsimplicity.spreadsheet.invalidFile" );
		expect( ()=>{
			var path = getTestFilePath( "test.tsv" );
			s.read( src=path );
		}).toThrow( type="cfsimplicity.spreadsheet.invalidFile" );
	})

	it( "the source file is in an old format not supported by POI", ()=>{
		expect( ()=>{
			var path = getTestFilePath( "oldformat.xls" );
			s.read( src=path );
		}).toThrow( type="cfsimplicity.spreadsheet.oldExcelFormatException" );
	})

})
</cfscript>