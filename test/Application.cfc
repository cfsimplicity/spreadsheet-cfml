component{
	this.name = "spreadsheet-cfml-tests";
	this.sessionManagement = false;
	this.applicationTimeout = CreateTimeSpan( 0, 0, 5, 0 );
	variables.relativePathToRoot = "../";
	this.mappings[ "/root" ] = GetDirectoryFromPath( GetCurrentTemplatePath() ) & relativePathToRoot;
	if( url.KeyExists( "nullSupport" ) )
		this.enableNullSupport = BooleanFormat( url.nullSupport ); //Testbox 6.3.1 throws exception in ACF if null support enabled
}