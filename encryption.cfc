component{

	/*
		I allow the POI JavaLoader instance to be maintained as the "contextLoader" and thereby used by POI objects when instantiating other POI classes. The EncryptionInfo class needs to load an EncryptionInfoBuilder class internally. For details of why this is so complex, see:
		https://github.com/markmandel/JavaLoader/wiki/Switching-the-ThreadContextClassLoader
	*/

	function init( required javaloader, required string algorithm ){
		variables.javaloader = javaloader;
		variables.switchThreadContextClassLoader = javaloader.switchThreadContextClassLoader;
		variables.algorithm = algorithm;
		return this;
	}

	function loadInfo(){
		var mode = javaloader.create( "org.apache.poi.poifs.crypt.EncryptionMode" );
		switch( algorithm ){
			case "agile":
				return javaloader.create( "org.apache.poi.poifs.crypt.EncryptionInfo" ).init( mode.agile );
			case "standard":
				return javaloader.create( "org.apache.poi.poifs.crypt.EncryptionInfo" ).init( mode.standard );
			case "binaryRC4":
				return javaloader.create( "org.apache.poi.poifs.crypt.EncryptionInfo" ).init( mode.binaryRC4 );
		}
	}

	function loadInfoWithSwitchedContextLoader(){
		return switchThreadContextClassLoader( "loadInfo", javaLoader.getURLClassLoader() );
	}

}