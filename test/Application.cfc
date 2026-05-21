component{
	this.name = "spreadsheet-cfml-tests";
	this.sessionManagement = false;
	this.applicationTimeout = CreateTimeSpan( 0, 0, 5, 0 );
	variables.relativePathToRoot = "../";
	this.mappings[ "/root" ] = GetDirectoryFromPath( GetCurrentTemplatePath() ) & relativePathToRoot;
}