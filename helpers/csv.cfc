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

}