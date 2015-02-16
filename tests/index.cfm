<cfscript>
paths = [ "root.tests.spreadsheet" ];
testRunner = New testbox.system.Testbox( paths );
echo( testRunner.run() );
</cfscript>