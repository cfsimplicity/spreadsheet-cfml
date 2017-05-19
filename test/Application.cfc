component{
	this.name = "luceeSpreadsheetTests";
	this.sessionManagement = false;
	this.applicationTimeout = CreateTimeSpan( 0, 0, 5, 0 );
	variables.relativePathToRoot = "../";
	this.mappings[ "/root" ] = GetDirectoryFromPath( GetCurrentTemplatePath() ) & relativePathToRoot;
}