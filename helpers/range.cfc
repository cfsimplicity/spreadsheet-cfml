component extends="base" accessors="true"{

	array function extractRanges( required string rangeList, required workbook, string dimension="row" ){
		/*
		Parses and validates a list of row/column numbers. Returns an array of structures with the keys: startAt, endAt
		@rangeList: a comma-delimited list where each value can be either a single number, a range of numbers with a hyphen, e.g. 1-5, or an open ended range, e.g. 2-. White space ignored.
		@dimension: "column" or "row"
		*/
		var result = [];
		var rangeTest = "^\d+(?:-\d*)?$";
		var ranges = ListToArray( arguments.rangeList );
		for( var thisRange in ranges ){
			thisRange = removeAllWhiteSpaceFrom( thisRange );
			if( !thisRange.REFind( rangeTest ) )
				Throw( type=library().getExceptionType(), message="Invalid range value", detail="The range value '#thisRange#' is not valid." );
			thisRange = handleOpenEndedRange( thisRange, arguments.dimension, arguments.workbook );
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

	private string function handleOpenEndedRange( required string range, required string dimension, required workbook ){
		var openEndedRangeTest = "^\d+-$";
		if( !arguments.range.REFind( openEndedRangeTest ) )
			return arguments.range;
		if( arguments.dimension == "column" )
			return arguments.range & library().getColumnCount( arguments.workbook );
		return arguments.range & library().getRowCount( arguments.workbook );
	}

}