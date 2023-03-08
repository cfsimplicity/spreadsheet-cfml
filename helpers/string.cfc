component extends="base"{

	any function newJavaStringBuilder(){
		return CreateObject( "Java", "java.lang.StringBuilder" ).init();
	}

}