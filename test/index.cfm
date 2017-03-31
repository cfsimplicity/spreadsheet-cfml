<cfscript>
paths = [ "root.test.suite" ];
try{
	testRunner = New testbox.system.TestBox( paths );
	WriteOutput( testRunner.run() );
}
catch( any exception ){
	WriteDump( exception );
}
</cfscript>