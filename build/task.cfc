component{

	/* Run me from CommandBox to compile source java, update the lib directory and build the lib-osgi.jar */

	void function run(){
		var rootPath = fileSystemUtil.resolvePath( "../" );
		var srcPath = rootPath & "src/";
		var libPath = rootPath & "lib/";
		var classpathDirectories = [
			srcPath
			,libPath & "poi-ooxml-5.0.0.jar"
			,libPath & "xmlbeans-4.0.0.jar"
		];
		var jarFileName = "luceeSpreadsheet.jar";
		var classNames = [ "HeaderImageVML" ]; //allows for more source files in future
		classNames.Each( function( className ){
			var classFileName = className & ".class";
			var javaSourceFilePath = srcPath & "luceeSpreadsheet/" & className & ".java";
			compileSource( classpathDirectories, srcPath, javaSourceFilePath );
		});
		createNewJar( jarFileName );
		replaceJarInLib( libPath, jarFileName );
		recreateOsgiJar( rootPath, libPath );
	}

	private void function compileSource( required array classpathDirectories, required string srcPath, required string javaSourceFilePath ){
		var destinationPath = getCWD() & "temp/";
		var args = "-sourcepath #arguments.srcPath# -classpath #arguments.classpathDirectories.ToList( ';' )# #arguments.javaSourceFilePath# -d #destinationPath#";
		execute name="javac" arguments=args timeout="5";
	}

	private void function createNewJar( required string jarFileName ){
		var tempSourcePath = getCWD() & "temp/";
		args = "--create --file #( getCWD() & arguments.jarFileName )# -C #tempSourcePath# .";//everything in temp
		execute name="jar" arguments=args timeout="5";
		DirectoryDelete( tempSourcePath, true );
	}

	private void function replaceJarInLib( required string libPath, required string jarFileName ){
		var tempFilePath =  getCWD() & jarFileName;
		var libJarPath = arguments.libPath & arguments.jarFileName;
		deleteFileIfExists( libJarPath );
		FileMove( tempFilePath, libJarPath );
		deleteFileIfExists( tempFilePath );
	}

	private void function recreateOsgiJar( required string rootPath, required string libPath ){
		var libOsgiPath = arguments.rootPath & "lib-osgi.jar";
		deleteFileIfExists( libOsgiPath );
		var manifestPath = getCWD() & "lib-osgi.mf";
		args = "--create --file #libOsgiPath# --manifest #manifestPath# -C #arguments.libPath# .";//everything in libPath
		execute name="jar" arguments=args timeout="5";
	}

	private void function deleteFileIfExists( required string path ){
		if( FileExists( arguments.path ) ) FileDelete( arguments.path );
	}

}