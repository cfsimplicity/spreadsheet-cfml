component extends="base"{

	any function getFormatObject( string type="DEFAULT" ){
		return library().createJavaObject( "org.apache.commons.csv.CSVFormat" )[ JavaCast( "string", arguments.type ) ];	
	}

	boolean function delimiterIsTab( required string delimiter ){
		return ArrayFindNoCase( [ "#Chr( 9 )#", "\t", "tab" ], arguments.delimiter );//CF2016 doesn't support [].FindNoCase( needle )
	}

	any function getFormat( required string delimiter ){
		if( arguments.delimiter.Len() )
			return getCsvFormatForDelimiter( arguments.delimiter );
		var format = getFormatObject( "RFC4180" );
		return format.builder()
				.setIgnoreSurroundingSpaces( JavaCast( "boolean", true ) )
				.build();
	}

	any function getFormatForPrinting( required string delimiter ){
		if( delimiterIsTab( arguments.delimiter ) )
			return getFormatObject( "TDF" );
		var format = getFormatObject( "EXCEL" );
		return format.builder().setDelimiter( arguments.delimiter ).build();
	}

	array function getColumnNames( required boolean firstRowIsHeader, required array data, required numeric maxColumnCount ){
		var result = [];
		if( arguments.firstRowIsHeader )
			var headerRow = arguments.data[ 1 ];
		for( var i=1; i <= arguments.maxColumnCount; i++ ){
			if( arguments.firstRowIsHeader && !IsNull( headerRow[ i ] ) && headerRow[ i ].Len() ){
				result.Append( headerRow[ i ] );
				continue;
			}
			result.Append( "column#i#" );
		}
		return result;
	}

	/* row order is not guaranteed if using more than one thread */
	array function queryToArrayForCsv( required query query, required boolean includeHeaderRow, numeric threads=1 ){
		var result = [];
		var columns = getQueryHelper()._QueryColumnArray( arguments.query );
		if( arguments.includeHeaderRow )
			result.Append( columns );
		if( ( arguments.threads > 1 ) && !library().engineSupportsParallelLoopProcessing() )
			getExceptionHelper().throwParallelOptionNotSupportedException();
		if( arguments.threads > 1 ){
			arguments.query.Each(
				function( row ){
					result.Append( getQueryRowValues( row, columns ) );
				}
				,true
				,arguments.threads
			);
			return result;			
		}
		for( var row IN arguments.query ){
			result.Append( getQueryRowValues( row, columns ) );
		}
		return result;
	}

	struct function parseFromString( required string csvString, required boolean trim, required any format ){
		if( arguments.trim )
			arguments.csvString = arguments.csvString.Trim();
		try{
			var parser = library().createJavaObject( "org.apache.commons.csv.CSVParser" ).parse( csvString, format );
			return dataFromParser( parser );
		}
		finally{
			getFileHelper().closeLocalFileOrStream( local, "parser" );
		}
	}

	struct function parseFromFile( required string path, required boolean trim, required any format ){
		getFileHelper()
			.throwErrorIFfileNotExists( arguments.path )
			.throwErrorIFnotCsvOrTextFile( arguments.path );
		return parseFromString( FileRead( arguments.path ), arguments.trim, arguments.format );
	}

	/* Private */
	private any function getCsvFormatForDelimiter( required string delimiter ){
		if( delimiterIsTab( arguments.delimiter ) )
			return getFormatObject( "TDF" );
		var format = getFormatObject( "RFC4180" );
		return format.builder()
			.setDelimiter( arguments.delimiter )
			.setIgnoreSurroundingSpaces( JavaCast( "boolean", true ) )//stop spaces between fields causing problems with embedded lines
			.build();
	}

	private struct function dataFromParser( required any parser ){
		var result = {
			data: []
			,maxColumnCount: 0
		};
		var recordIterator = arguments.parser.iterator();
		while( recordIterator.hasNext() ){
			var record = recordIterator.next();
			result.maxColumnCount = Max( result.maxColumnCount, record.size() );
			//ACF can't handle native arrays in QueryNew()
			var values = ( library().getIsACF() )? record.toList(): record.values();
			ArrayAppend( result.data, values );
		}
		return result;
	}

	private array function getQueryRowValues( required row, required array columns ){
		var rowValues = [];
		for( var column IN arguments.columns ){
			var cellValue = arguments.row[ column ];
			if( getDateHelper().isDateObject( cellValue ) )
				cellValue = DateTimeFormat( cellValue, library().getDateFormats().DATETIME );
			if( IsValid( "integer", cellValue ) )
				cellValue = JavaCast( "string", cellValue );// prevent CSV writer converting 1 to 1.0
			rowValues.Append( cellValue );
		};
		return rowValues;
	}

}