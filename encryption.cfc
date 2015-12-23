component{

	/*
		I allow the POI JavaLoader instance to be maintained as the "contextLoader" and thereby used by POI objects when instantiating other POI classes. The EncryptionInfo class needs to load an EncryptionInfoBuilder class internally. For details of why this is so complex, see:
		https://github.com/markmandel/JavaLoader/wiki/Switching-the-ThreadContextClassLoader
	*/

	function init( required javaloader,required string algorithm ){
		variables.javaloader=javaloader;
		variables.switchThreadContextClassLoader=javaloader.switchThreadContextClassLoader;
		variables.algorithm=algorithm;
		return this;
	}

	function loadInfo(){
		var mode=javaloader.create( "org.apache.poi.poifs.crypt.EncryptionMode" );
		switch( algorithm ){
			case "agile":
				var info=javaloader.create( "org.apache.poi.poifs.crypt.EncryptionInfo" ).init( mode.agile );
				break;
			case "standard":
				var info=javaloader.create( "org.apache.poi.poifs.crypt.EncryptionInfo" ).init( mode.standard );
				break;
			case "binaryRC4":
				var info=javaloader.create( "org.apache.poi.poifs.crypt.EncryptionInfo" ).init( mode.binaryRC4 );
				break;
		}
		return info;
	}

	function loadInfoWithSwitchedContextLoader(){
		return switchThreadContextClassLoader( "loadInfo",javaLoader.getURLClassLoader() );
	}

}