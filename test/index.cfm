<cfscript>
paths = [ "root.test.suite" ];
try{
	testRunner = New testbox.system.TestBox( paths );
	WriteOutput( testRunner.run(reporter=url?.reporter ?: 'simple') );
}
catch( any exception ){
	WriteDump( exception );
}
</cfscript>