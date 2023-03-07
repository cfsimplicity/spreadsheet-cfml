component extends="base"{

	void function dumpPathToClassNoOsgi( required string className ){
		var classLoader = loadClass( arguments.className ).getClass().getClassLoader();
		var path = classLoader.getResource( arguments.className.Replace( ".", "/", "all" ) & ".class" ).getPath();
		WriteDump( path );
	}

	any function loadClass( required string javaclass ){
		if( library().getRequiresJavaLoader() )
			return loadClassUsingJavaLoader( arguments.javaclass );
		if( !IsNull( library().getOsgiLoader() ) )
			return loadClassUsingOsgi( arguments.javaclass );
		// If ACF and not using JL, *the correct* POI jars must be in the class path and any older versions *removed*
		try{
			library().setJavaClassesLastLoadedVia( "The java class path" );
			return CreateObject( "java", arguments.javaclass );
		}
		catch( any exception ){
			return loadClassUsingJavaLoader( arguments.javaclass );
		}
	}

	/* Private */

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

}