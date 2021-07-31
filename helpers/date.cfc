component extends="base" accessors="true"{

	property name="dateUtil" getter="false" setter="false";

	public any function getDateUtil(){
		if( IsNull( variables.dateUtil ) )
			variables.dateUtil = getClassHelper().loadClass( "org.apache.poi.ss.usermodel.DateUtil" );
		return variables.dateUtil;
	}

	public any function setFormats( struct dateFormats ){
		library().setDateFormats( defaultFormats() );
		if( !arguments.KeyExists( "dateFormats" ) )
			return this;
		var libraryInstanceFormats = library().getDateFormats();
		for( var format in arguments.dateFormats ){
			if( !libraryInstanceFormats.KeyExists( format ) )
				Throw( type=library().getExceptionType(), message="Invalid date format key", detail="'#format#' is not a valid dateformat key. Valid keys are DATE, DATETIME, TIME and TIMESTAMP" );
			libraryInstanceFormats[ format ] = arguments.dateFormats[ format ];
		}
		return this;
	}

	public string function getDefaultDateMaskFor( required date value ){
		var dateOnly = CreateDate( Year( arguments.value ), Month( arguments.value ), Day( arguments.value ) );
		if( isDateOnlyValue( arguments.value ) )
			return library().getDateFormats().DATE;
		if( isTimeOnlyValue( arguments.value ) )
			return library().getDateFormats().TIME;
		return library().getDateFormats().TIMESTAMP;
	}

	public boolean function isDateObject( required input ){
		return IsInstanceOf( arguments.input, "java.util.Date" );
	}

	public boolean function isDateOnlyValue( required date value ){
		var dateOnly = CreateDate( Year( arguments.value ), Month( arguments.value ), Day( arguments.value ) );
		return ( DateCompare( arguments.value, dateOnly, "s" ) == 0 );
	}

	public boolean function isTimeOnlyValue( required date value ){
		//NB: this will only detect CF time object (epoch = 1899-12-30), not those using unix epoch 1970-01-01
		return ( Year( arguments.value ) == "1899" );
	}

	/* Private */
	
	private struct function defaultFormats(){
		return {
			DATE: "yyyy-mm-dd"
			,DATETIME: "yyyy-mm-dd HH:nn:ss"
			,TIME: "hh:mm:ss"
			,TIMESTAMP: "yyyy-mm-dd hh:mm:ss"
		};
	}

}