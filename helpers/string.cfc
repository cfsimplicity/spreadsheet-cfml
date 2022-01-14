component extends="base" accessors="true"{

	any function newJavaStringBuilder(){
		return CreateObject( "Java", "java.lang.StringBuilder" ).init();
	}

}