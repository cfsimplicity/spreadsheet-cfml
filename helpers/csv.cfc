component extends="base" accessors="true"{

	public boolean function delimiterIsTab( required string delimiter ){
		return ArrayFindNoCase( [ "#Chr( 9 )#", "\t", "tab" ], arguments.delimiter );//CF2016 doesn't support [].FindNoCase( needle )
	}

	public any function getCsvFormatForDelimiter( required string delimiter ){
		if( delimiterIsTab( arguments.delimiter ) )
			return getClassHelper().loadClass( "org.apache.commons.csv.CSVFormat" )[ JavaCast( "string", "TDF" ) ];
		return getClassHelper().loadClass( "org.apache.commons.csv.CSVFormat" )[ JavaCast( "string", "RFC4180" ) ]
			.withDelimiter( JavaCast( "char", arguments.delimiter ) )
			.withIgnoreSurroundingSpaces();//stop spaces between fields causing problems with embedded lines
	}

	public string function readFile( required string filepath ){
		getFileHelper()
			.throwErrorIFfileNotExists( arguments.filepath )
			.throwErrorIFnotCsvOrTextFile( arguments.filepath );
		return FileRead( arguments.filepath );
	}

	public struct function dataFromRecords( required array records ){
		var result = {
			data: []
			,maxColumnCount: 0
		};
		for( var record in arguments.records ){
			var row = [];
			var columnNumber = 0;
			var iterator = record.iterator();
			while( iterator.hasNext() ){
				columnNumber++;
				result.maxColumnCount = Max( result.maxColumnCount, columnNumber );
				row.Append( iterator.next() );
			}
			result.data.Append( row );
		}
		return result;
	}

	public array function getColumnNames( required boolean firstRowIsHeader, required array data, required numeric maxColumnCount ){
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

}