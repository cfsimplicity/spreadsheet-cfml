component{

	/*
		I allow the POI JavaLoader instance to be maintained as the "contextLoader" and thereby used by POI objects when instantiating other POI classes. The EncryptionInfo class needs to load an EncryptionInfoBuilder class internally. For details of why this is so complex, see:
		https://github.com/markmandel/JavaLoader/wiki/Switching-the-ThreadContextClassLoader
	*/

	function init( required javaloader ){
		variables.javaloader=javaloader;
		variables.switchThreadContextClassLoader=javaloader.switchThreadContextClassLoader;
		return this;
	}

	function loadInfo(){
		var mode=javaloader.create( "org.apache.poi.poifs.crypt.EncryptionMode" );
		//NB:  Excel viewer doesn't seem to be able to open the file if agile and standard modes are used
		var info=javaloader.create( "org.apache.poi.poifs.crypt.EncryptionInfo" ).init( mode.binaryRC4 );
		return info;
	}

	function loadInfoWithSwitchedContextLoader(){
		return switchThreadContextClassLoader( "loadInfo",javaLoader.getURLClassLoader() );
	}

}