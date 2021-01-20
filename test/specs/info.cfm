<cfscript>
describe( "info", function(){

	beforeEach( function(){
		variables.infoToAdd = {
			author: "Bob"
			,category: "Testing"
			,lastAuthor: "Anne"
			,comments: "OK"
			,keywords: "test"
			,manager: "Diane"
			,company: "Acme Ltd"
			,subject: "tests"
			,title: "Test figures"
		};
		var additional = {
			creationDate: DateFormat( Now(), "yyyymmdd" )
			,lastEdited: ""
			,lastSaved: ""
			,sheetnames: "Sheet1"
			,sheets: 1
			,spreadSheetType: "Excel"
		};
		variables.infoToBeReturned = Duplicate( infoToAdd );
		StructAppend( infoToBeReturned, additional );
	});

	it( "Adds and can get back info from a binary xls", function(){
		var workbook = s.new();
		s.addInfo( workbook, infoToAdd );
		var expected = infoToBeReturned;
		var actual = s.info( workbook );
		actual.creationDate = DateFormat( Now(), "yyyymmdd" );// Doesn't return this value so mock
		expect( actual ).toBe( expected );
	});

	it( "Adds and can get back info from an xml xlsx", function(){
		var workbook = s.newXlsx();
		s.addInfo( workbook, infoToAdd );
		infoToBeReturned.spreadSheetType = "Excel (2007)";
		expected = infoToBeReturned;
		var actual = s.info( workbook );
		actual.creationDate = DateFormat( actual.creationDate, "yyyymmdd" ); // Doesn't return this value so mock
		expect( actual ).toBe( expected );
	});

	it( "Adds and can get back info from a streaming xlsx", function(){
		var workbook = s.newStreamingXlsx();
		s.addInfo( workbook, infoToAdd );
		infoToBeReturned.spreadSheetType = "Excel (2007)";
		var expected = infoToBeReturned;
		var actual = s.info( workbook );
		actual.creationDate = DateFormat( actual.creationDate, "yyyymmdd" ); // Doesn't return this value so mock
		expect( actual ).toBe( expected );
	});

	it( "Handles missing lastAuthor value in an xlsx", function(){
		infoToAdd.delete( "lastAuthor" );
		infoToBeReturned.delete( "lastAuthor" );
		var workbook = s.newXlsx();
		s.addInfo( workbook, infoToAdd );
		infoToBeReturned.spreadSheetType = "Excel (2007)";
		var expected = infoToBeReturned;
		var actual = s.info( workbook );
		actual.creationDate = DateFormat( actual.creationDate, "yyyymmdd" ); // Doesn't return this value so mock
		expect( actual ).toBe( expected );
	});

	it( "Can accept a file path instead of a workbook", function(){
		var workbook = s.new();
		s.addInfo( workbook, infoToAdd );
		s.write( workbook, tempXlsPath, true );
		var expected = infoToBeReturned;
		var actual = s.info( tempXlsPath );
		actual.creationDate = DateFormat( Now(), "yyyymmdd" );// Doesn't return this value so mock
		expect( actual ).toBe( expected );
	});

	afterEach( function(){
		if( FileExists( variables.tempXlsPath ) ) FileDelete( variables.tempXlsPath );
		if( FileExists( variables.tempXlsxPath ) ) FileDelete( variables.tempXlsxPath );
	});

});	
</cfscript>
