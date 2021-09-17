component extends="base" accessors="true"{

	public any function newJavaStringBuilder(){
		return CreateObject( "Java", "java.lang.StringBuilder" ).init();
	}

}