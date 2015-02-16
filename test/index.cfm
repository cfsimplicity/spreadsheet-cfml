<cfscript>
paths = [ "root.test.suite" ];
testRunner = New testbox.system.Testbox( paths );
echo( testRunner.run() );
</cfscript>