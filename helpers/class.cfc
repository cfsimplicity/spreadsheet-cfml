component extends="base"{

	// common class name references
	string function getClassName( required string objectName ){
		switch( arguments.objectName ){
			case "HSSFWorkbook": return "org.apache.poi.hssf.usermodel.HSSFWorkbook";
			case "XSSFWorkbook": return "org.apache.poi.xssf.usermodel.XSSFWorkbook";
			case "SXSSFWorkbook": return "org.apache.poi.xssf.streaming.SXSSFWorkbook";
		}
	}

	void function dumpPathToClassNoOsgi( required string className ){
		var classLoader = loadClass( arguments.className ).getClass().getClassLoader();
		var path = classLoader.getResource( arguments.className.Replace( ".", "/", "all" ) & ".class" ).getPath();
		WriteDump( path );
	}

	any function loadClass( required string javaclass ){
		switch( library().getLoadJavaClassesUsing() ){
			case "JavaLoader": return loadClassUsingJavaLoader( arguments.javaclass );
			case "osgi": return loadClassUsingOsgi( arguments.javaclass );
			case "dynamicPath": return loadClassUsingDynamicPath( arguments.javaclass );
		}
		//classPath or app's javaSettings
		var lastLoadedVia = ( library().getLoadJavaClassesUsing() == "javaSettings" )? "Application javaSettings": "The java class path";
		library().setJavaClassesLastLoadedVia( lastLoadedVia );
		return CreateObject( "java", arguments.javaclass );
	}

	string function validateLoadingMethod( required string method ){
		if( isValidLoadingMethod( arguments.method ) )
			return arguments.method;
		Throw( type=library().getExceptionType() & ".invalidJavaLoadingMethod", message="'#arguments.method#' is not valid. Valid methods are #validLoadingMethods().ToList( ', ' )#" );
	}

	/* Private */

	private array function validLoadingMethods(){
		return [ "osgi", "JavaLoader", "dynamicPath", "classPath", "javaSettings" ];
	}

	private boolean function isValidLoadingMethod( required string method ){
		return validLoadingMethods().FindNoCase( arguments.method );
	}

	private any function loadClassUsingJavaLoader( required string javaclass ){
		library().setJavaClassesLastLoadedVia( "JavaLoader" );
		return library().getJavaLoaderInstance().create( arguments.javaclass );
	}

	private any function loadClassUsingOsgi( required string javaclass ){
		library().setJavaClassesLastLoadedVia( "OSGi bundle" );
		try{
			return library().getOsgiLoader().loadClass(
				className: arguments.javaclass
				,bundlePath: variables.rootPath & "lib-osgi.jar"
				,bundleSymbolicName: library().getOsgiLibBundleSymbolicName()
				,bundleVersion: library().getOsgiLibBundleVersion()
			);
		}
		catch( org.osgi.framework.BundleException exception ){
			if( exception.message.FindNoCase( "is not available in version" ) ){
				library().flushOsgiBundle();
				retry;
			}
			else
				rethrow;
		}
	}

	private any function loadClassUsingDynamicPath( required string javaclass ){
		library().setJavaClassesLastLoadedVia( "Dynamic path" );
		return CreateObject( "java", arguments.javaclass, library().getLibPath() );
	}

}