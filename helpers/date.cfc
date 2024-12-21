component extends="base"{

	property name="dateUtil";

	any function getDateUtil(){
		if( IsNull( variables.dateUtil ) )
			variables.dateUtil = library().createJavaObject( "org.apache.poi.ss.usermodel.DateUtil" );
		return variables.dateUtil;
	}

	struct function defaultFormats(){
		return {
			DATE: "yyyy-mm-dd"
			,DATETIME: "yyyy-mm-dd HH:nn:ss"
			,TIME: "hh:mm:ss"
			,TIMESTAMP: "yyyy-mm-dd hh:mm:ss"
		};
	}

	any function setCustomFormats( required struct dateFormats ){
		var libraryInstanceFormats = library().getDateFormats();
		for( var format in arguments.dateFormats ){
			if( !libraryInstanceFormats.KeyExists( format ) )
				Throw( type=library().getExceptionType() & ".invalidDateFormatKey", message="Invalid date format key", detail="'#format#' is not a valid dateformat key. Valid keys are DATE, DATETIME, TIME and TIMESTAMP" );
			libraryInstanceFormats[ format ] = arguments.dateFormats[ format ];
		}
		return this;
	}

	string function getDefaultDateMaskFor( required date value ){
		if( isDateOnlyValue( arguments.value ) )
			return library().getDateFormats().DATE;
		if( isTimeOnlyValue( arguments.value ) )
			return library().getDateFormats().TIME;
		return library().getDateFormats().TIMESTAMP;
	}

	boolean function isDateObject( required input ){
		return IsInstanceOf( arguments.input, "java.util.Date" );
	}

	boolean function isDateOnlyValue( required date value ){
		var dateOnly = CreateDate( Year( arguments.value ), Month( arguments.value ), Day( arguments.value ) );
		return ( DateCompare( arguments.value, dateOnly, "s" ) == 0 );
	}

	boolean function isTimeOnlyValue( required date value ){
		//NB: this will only detect CF time object (epoch = 1899-12-30), not those using unix epoch 1970-01-01
		return ( Year( arguments.value ) == "1899" );
	}

	string function getPoiTimeZone(){
		return library().createJavaObject( "org.apache.poi.util.LocaleUtil" ).getUserTimeZone().toString();
	}

	any function matchPoiTimeZoneToEngine(){
		//ACF doesn't allow the server/context timezone to be changed
		//Boxlang supports setting the context timezone, but not getting it?
		if( !library().getIsLucee() )
			return this;
		//Lucee allows the context timezone to be changed, which can cause problems with date calculations
		if( getPoiTimeZone() == GetTimezone() )
			return this;
		//Make POI match the Lucee timezone for the duration of the current thread
		library().createJavaObject( "org.apache.poi.util.LocaleUtil" ).setUserTimeZone( GetTimezone() );
		return this;
	}

	any function convertDateToJava( required date object ){
		return CreateObject( "java", "java.util.Date" ).init( arguments.object.getTime() );
	}

	// alternative BIFS
	boolean function _IsDate( required value ){
		if( !IsDate( arguments.value ) )
			return false;
		// Lucee will treat 01-23112 or 23112-01 as a date!
		if( ParseDateTime( arguments.value ).Year() > 9999 ) //ACF future limit
			return false;
		// ACF accepts "9a", "9p", "9 a" as dates
		// ACF no member function
		if( REFind( "^\d+\s*[apAP]{1,1}$", arguments.value ) )
			return false;
		return true;
	}

	any function _ParseDateTime( required value ){
		// ACF and boxlang can test for an existing date object
		if( !library().getIsLucee() && IsDateObject( arguments.value ) )
			return arguments.value;
		//Boxlang is very limited in what it will accept
		if( !library().getIsBoxlang() )
			return ParseDateTime( arguments.value );
		//Boxlang: support a limited set of "non-standard" date/time formats
		if( arguments.value.REFindNoCase( "^\d{1,2}\D\d{4,4}$" ) ){ //e.g. 01.2024 or 04/2024
			arguments.value = arguments.value.REReplaceNoCase( "(\d{2,2})\D(\d{4,4})", "01/\1/\2" );
			return ParseDateTime( arguments.value );
		}
		if( arguments.value.REFindNoCase( "^\d{2,2}:\d{2,2}$" ) ) //e.g. //08:21
			return ParseDateTime( "1899-12-30T#arguments.value#:00Z" );
		if( arguments.value.REFindNoCase( "^\d{2,2}:\d{2,2}:\d{2,2}$" ) ) //e.g. //08:21:30
			return ParseDateTime( "1899-12-30T#arguments.value#Z" );
		return ParseDateTime( arguments.value );
	}

}