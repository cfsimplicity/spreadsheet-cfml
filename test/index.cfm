<cfscript>
paths = [ "root.test.suite" ];
testRunner = New testbox.system.TestBox( paths );
echo( testRunner.run() );
</cfscript>