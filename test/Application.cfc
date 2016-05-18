component{
	this.name	=	"luceeSpreadSheetTests";
	this.sessionManagement	=	false;
	this.applicationTimeout = CreateTimeSpan( 0,0,5,0 );
	request.relativePathToRoot	=	"../"; // from this directory to the topmost application directory
	this.mappings[ "/root" ]		=	GetDirectoryFromPath( GetCurrentTemplatePath() ) & request.relativePathToRoot;
	this.mappings[ "/mxunit" ] 	= ExpandPath( "/testbox/system/compat/" );
}