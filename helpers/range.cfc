component extends="base"{

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
				Throw( type=library().getExceptionType() & ".invalidRange", message="Invalid range value", detail="The range value '#thisRange#' is not valid." );
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

	any function getCellRangeAddressFromColumnAndRowIndices( required struct indices ){
		//index = 0 based
		return library().createJavaObject( "org.apache.poi.ss.util.CellRangeAddress" ).init(
			JavaCast( "int", arguments.indices.startRow )
			,JavaCast( "int", arguments.indices.endRow )
			,JavaCast( "int", arguments.indices.startColumn )
			,JavaCast( "int", arguments.indices.endColumn )
		);
	}

	any function getCellRangeAddressFromRowIndex( required workbook, required numeric rowIndex ){
		var indices = {
			startRow: arguments.rowIndex
			,endRow: arguments.rowIndex
			,startColumn: 0
			,endColumn: ( library().getColumnCount( arguments.workbook ) -1 )
		};
		return getCellRangeAddressFromColumnAndRowIndices( indices );
	}

	any function getCellRangeAddressFromReference( required string rangeReference ){
		/*
		rangeReference = usually a standard area ref (e.g. "B1:D8"). May be a single cell ref (e.g. "B5") in which case the result is a 1 x 1 cell range. May also be a whole row range (e.g. "3:5"), or a whole column range (e.g. "C:F")
		*/
		return library().createJavaObject( "org.apache.poi.ss.util.CellRangeAddress" ).valueOf( JavaCast( "String", arguments.rangeReference ) );
	}

	string function convertRangeReferenceToAbsoluteAddress( required string rangeReference ){
		return arguments.rangeReference.ReplaceAll( "([A-Za-z]+|\d+)", "\$$1" ).UCase(); //Use java regex for group reference consistency
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