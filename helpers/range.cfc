component extends="base" accessors="true"{

	public array function extractRanges( required string rangeList ){
		/*
		A range is a comma-delimited list of ranges, and each value can be either a single number or a range of numbers with a hyphen. Ignores any white space.
		Parses and validates a list of row/column numbers. Returns an array of structures with the keys: startAt, endAt
		*/
		var result = [];
		var rangeTest = "^[0-9]{1,}(-[0-9]{1,})?$";
		var ranges = ListToArray( arguments.rangeList );
		for( var thisRange in ranges ){
			thisRange = removeAllWhiteSpaceFrom( thisRange );
			if( !REFind( rangeTest, thisRange ) )
				Throw( type=library().getExceptionType(), message="Invalid range value", detail="The range value '#thisRange#' is not valid." );
			var parts = ListToArray( thisRange, "-" );
			//if this is a single number, the start/endAt values are the same
			var range = {
				startAt: parts[ 1 ]
				,endAt: parts[ parts.Len() ]
			};
			result.Append( range );
		}
		return result;
	}

	/* Private */
	private string function removeAllWhiteSpaceFrom( required string value ){
		return arguments.value.REReplace( "\s+", "", "ALL" );
	}

}