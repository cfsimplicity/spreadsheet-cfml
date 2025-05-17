<cfscript>
describe(
  title="Lucee only timezone tests",
  body=function(){

    it( "Doesn't offset a date value even if the Lucee timezone doesn't match the system", ()=>{
      variables.currentTZ = GetTimeZone();
      variables.tempTZ = "US/Eastern";
      spreadsheetTypes.Each( ( type )=>{
        SetTimeZone( tempTZ );
        var path = variables[ "temp" & type & "Path" ];
        local.s = newSpreadsheetInstance();//timezone mismatch detection cached is per instance
        local.s.newChainable( type ).setCellValue( "2022-01-01", 1, 1, "date" ).write( path, true );
        var actual = local.s.read( path, "query" ).column1;
        var expected = CreateDate( 2022, 01, 01 );
        expect( actual ).toBe( expected );
        SetTimeZone( currentTZ );
      })

    })

  },
  skip=( !s.getIsLucee() || ( s.getDateHelper().getPoiTimeZone() != "Europe/London" ) )// only valid if system timezone is ahead of temporary test timezone
);
</cfscript>