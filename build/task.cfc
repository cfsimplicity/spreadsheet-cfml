component{

	/* Run me from CommandBox to compile source java, update the lib directory and build the lib-osgi.jar */

	property name="classpathDirectories";
	property name="libPath";
	property name="minimumSupportedJvmVersion" default="8";//update this as necessary
	property name="rootPath";
	property name="srcPath";
	property name="tempDirectoryPath";

	void function run(){
		variables.tempDirectoryPath = getCWD() & "temp/";
		if( !DirectoryExists( tempDirectoryPath ) )
			DirectoryCreate( tempDirectoryPath );
		variables.rootPath = fileSystemUtil.resolvePath( "../" );
		variables.srcPath = variables.rootPath & "src/";
		variables.libPath = variables.rootPath & "lib/";
		variables.classpathDirectories = [
			variables.srcPath
			,variables.libPath & "poi-ooxml-5.0.0.jar"
			,variables.libPath & "xmlbeans-4.0.0.jar"
		];
		var jarFileName = "spreadsheet-cfml.jar";
		var classNames = [ "HeaderImageVML" ]; //allows for more source files in future
		classNames.Each( function( className ){
			var classFileName = className & ".class";
			var javaSourceFilePath = variables.srcPath & "spreadsheetCFML/" & className & ".java";
			compileSource( javaSourceFilePath );
		});
		createNewJar( jarFileName );
		replaceJarInLib( jarFileName );
		recreateOsgiJar();
		if( DirectoryExists( tempDirectoryPath ) )
			DirectoryDelete( tempDirectoryPath, true );
	}

	private void function compileSource( required string javaSourceFilePath ){
		var destinationPath = variables.tempDirectoryPath;
		var args = "--release #variables.minimumSupportedJvmVersion# -sourcepath #variables.srcPath# -classpath #variables.classpathDirectories.ToList( ';' )# #arguments.javaSourceFilePath# -d #destinationPath#";
		execute name="javac" arguments=args timeout="5";
	}

	private void function createNewJar( required string jarFileName ){
		var tempSourcePath = variables.tempDirectoryPath;
		args = "cfM #( getCWD() & arguments.jarFileName )# -C #tempSourcePath# .";//everything in temp, M=no manifest to avoid JVM version problems 
		execute name="jar" arguments=args timeout="5";
	}

	private void function replaceJarInLib( required string jarFileName ){
		var tempFilePath = getCWD() & jarFileName;
		var libJarPath = variables.libPath & arguments.jarFileName;
		deleteFileIfExists( libJarPath );
		FileMove( tempFilePath, libJarPath );
		deleteFileIfExists( tempFilePath );
	}

	private void function recreateOsgiJar(){
		var libOsgiPath = variables.rootPath & "lib-osgi.jar";
		deleteFileIfExists( libOsgiPath );
		var manifestPath = getCWD() & "lib-osgi.mf";
		args = "cfm #libOsgiPath# #manifestPath# -C #variables.libPath# .";//everything in libPath
		execute name="jar" arguments=args timeout="5";
	}

	private void function deleteFileIfExists( required string path ){
		if( FileExists( arguments.path ) ) FileDelete( arguments.path );
	}

}