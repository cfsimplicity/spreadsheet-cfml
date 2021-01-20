component accessors="true"{

	/*
		I encapsulate methods for loading OSGi bundles dynamically
	*/

	property name="CFMLEngineFactory";
	property name="luceeOSGiUtil";
	property name="version" default="1.0.0" setter="false";

	public any function init(){
		variables.luceeOSGiUtil = CreateObject( "java", "lucee.runtime.osgi.OSGiUtil" );
		variables.CFMLEngineFactory = CreateObject( "java", "lucee.loader.engine.CFMLEngineFactory" ).getInstance();
		return this;
	}

	public any function loadClass( required string className, required string bundlePath, required string bundleSymbolicName, required string bundleVersion ){
		if( !bundleIsLoaded( arguments.bundleSymbolicName, arguments.bundleVersion ) ) installBundle( arguments.bundlePath );
		return CreateObject( "java", arguments.className, arguments.bundleSymbolicName, arguments.bundleVersion );
	}

	public void function installBundle( required string path ){
		var CFMLEngineFactory = this.getCFMLEngineFactory();
		var resource = CFMLEngineFactory.getResourceUtil().toResourceExisting( GetPageContext(), JavaCast( "string", arguments.path ) );
		this.getLuceeOSGiUtil().installBundle( CFMLEngineFactory.getBundleContext(), resource, JavaCast( "boolean", true ) );
	}

	public void function uninstallBundle( required string bundleSymbolicName, required string bundleVersion ){
		var bundle = getBundle( arguments.bundleSymbolicName, arguments.bundleVersion );
		if( IsNull( bundle ) ) return;
		bundle.uninstall();
	}

	public any function getBundle( required string bundleSymbolicName, required string bundleVersion ){
		var OSGiUtil = this.getLuceeOSGiUtil();
		return OSGiUtil.getBundleLoaded( arguments.bundleSymbolicName, OSGiUtil.toVersion( arguments.bundleVersion ), JavaCast( "null", "" ) );
	}

	public boolean function bundleIsLoaded( required string bundleSymbolicName, required string bundleVersion ){
		return !IsNull( getBundle( arguments.bundleSymbolicName, arguments.bundleVersion ) );
	}

}