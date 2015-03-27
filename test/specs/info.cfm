<cfscript>
describe( "info tests",function(){

	beforeEach( function(){
		variables.infoToAdd = {
			author="Bob"
			,category="Testing"
			,lastAuthor="Anne"
			,comments="OK"
			,keywords="test"
			,manager="Diane"
			,company="Acme Ltd"
			,subject="tests"
			,title="Test figures"
		};
		var additional = {
			creationDate=DateFormat( Now(),"yyyymmdd" )
			,lastEdited = ""
			,lastSaved = ""
			,sheetnames = "Sheet1"
			,sheets=1
			,spreadSheetType="Excel"
		};
		variables.infoToBeReturned = infoToAdd.Append( additional );
	});

	it( "Adds and can get back info from a binary xls",function() {
		workbook = s.new();
		s.addInfo( workbook,infoToAdd );
		expected = infoToBeReturned;
		actual = s.info( workbook );
		actual.creationDate=DateFormat( Now(),"yyyymmdd" );// Doesn't return this value so mock
		expect( actual ).toBe( expected );
	});

	it( "Adds and can get back info from an xml xlsx",function() {
		workbook = s.new( xmlformat=true );
		s.addInfo( workbook,infoToAdd );
		infoToBeReturned.spreadSheetType = "Excel (2007)";
		expected = infoToBeReturned;
		actual = s.info( workbook );
		actual.creationDate = DateFormat( actual.creationDate,"yyyymmdd" ); //can't test time obviously.
		expect( actual ).toBe( expected );
	});

});	
</cfscript>
