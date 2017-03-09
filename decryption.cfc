component{

	/*
		I allow the POI JavaLoader instance to be maintained as the "contextLoader" and thereby used by POI objects when instantiating other POI classes. The EncryptionInfo class needs to load an EncryptionInfoBuilder class internally. For details of why this is so complex, see:
		https://github.com/markmandel/JavaLoader/wiki/Switching-the-ThreadContextClassLoader
	*/

	function init( required javaloader, required poiFilesystem ){
		variables.javaloader = javaloader;
		variables.switchThreadContextClassLoader = javaloader.switchThreadContextClassLoader;
		variables.poiFilesystem = poiFilesystem;
		return this;
	}

	function loadInfo(){
		return javaloader.create( "org.apache.poi.poifs.crypt.EncryptionInfo" ).init( poiFilesystem );
	}

	function loadInfoWithSwitchedContextLoader(){
		return switchThreadContextClassLoader( "loadInfo", javaLoader.getURLClassLoader() );
	}

}