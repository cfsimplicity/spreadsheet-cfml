component extends="BaseCsv" accessors="true"{

	property name="data" setter="false";
	property name="parallelThreadsToUse" type="numeric" default=0 setter="false";
	property name="useQueryColumnsAsHeader" type="boolean" default="false" setter="false";
	property name="useStructKeysAsHeader" type="boolean" default="false" setter="false";

	public WriteCsv function init( required spreadsheetLibrary ){
		super.init( spreadsheetLibrary=arguments.spreadsheetLibrary, initialPredefinedFormat="EXCEL" );
		return this;
	}

	/* Public builder API */

	public WriteCsv function fromData( required any data ){
		if( !IsArray( arguments.data ) && !IsQuery( arguments.data ) )
			Throw( type=variables.library.getExceptionType() & ".invalidDataForCsv", message="Invalid data", detail="Please pass your data as a query, an array of arrays, or an array of structs" );
		variables.data = arguments.data;
		return this;
	}

	public WriteCsv function toFile( required string path ){
		if( arguments.path.Left( 4 ) == "ram:" )
			Throw( type=variables.library.getExceptionType() & ".vfsNotSupported", message="Invalid file path", detail="Virtual File System (RAM) paths are not supported for writing to CSV. Try just returning the CSV string and then using FileWrite() to write to the VFS path" );
		variables.filepath = arguments.path;
		return this;
	}

	public WriteCsv function withParallelThreads( numeric numberOfThreads=2 ){
		/* WARNING: can have unexpected results such as rows out of order or system crashes. USE WITH CARE. */
		if( arguments.numberOfThreads < 2 ){
			variables.parallelThreadsToUse = 0;
			return this;
		}
		variables.parallelThreadsToUse = Int( arguments.numberOfThreads );
		return this;
	}

	public WriteCsv function withQueryColumnsAsHeader( boolean state=true ){
		variables.useQueryColumnsAsHeader = arguments.state;
		return this;
	}

	public WriteCsv function withStructKeysAsHeader( boolean state=true ){
		variables.useStructKeysAsHeader = arguments.state;
		return this;
	}

	// final execution
	public any function execute(){
		if( IsNull( variables.data ) )
			Throw( type=variables.library.getExceptionType() & ".missingDataForCsv", message="Missing data", detail="Please specify the data you want to write using '.fromData( data )'" );
		var appendable = newAppendableBuffer();
		printTo( appendable );
		if( IsNull( variables.filepath ) )
			return appendable.toString();
		return this;
	}

	/* Private */
	private void function printTo( required appendable ){
		try{
			if( IsQuery( data ) ){
				setQueryColumnsAsHeaderIfRequired();
				var printer = newPrinter( arguments.appendable );
				printFromQuery( printer );
				return;
			}
			setStructKeysAsHeaderIfRequired();
			var printer = newPrinter( arguments.appendable );
			printFromArray( printer );
		}
		finally{
			if( local.KeyExists( "printer" ) )
				printer.close( JavaCast( "boolean", true ) );
		}
	}

	private void function setQueryColumnsAsHeaderIfRequired(){
		if( !variables.useQueryColumnsAsHeader )
			return;
		var columns = variables.library.getQueryHelper()._QueryColumnArray( variables.data );
		super.withHeader( columns );
	}

	private void function setStructKeysAsHeaderIfRequired(){
		if( !variables.useStructKeysAsHeader || !variables.data.Len() || !IsStruct( variables.data[ 1 ] ) )
			return;
		var keys = variables.data[ 1 ].KeyArray();
		super.withHeader( keys );
	}

	private any function newPrinter( required appendable ){
		return variables.library.createJavaObject( "org.apache.commons.csv.CSVPrinter" ).init( arguments.appendable, variables.format );
	}

	private any function newAppendableBuffer(){
		if( IsNull( variables.filepath ) )
			return variables.library.getStringHelper().newJavaStringBuilder();
		return newBufferedFileWriter();
	}

	private any function newBufferedFileWriter(){
		var charset = CreateObject( "java", "java.nio.charset.Charset" ).forName( "UTF-8" );
		if( variables.library.getIsBoxlang() ){
			//boxlang (or java21?) doesn't recognize ACF/Lucee signatures
			var path = CreateObject( "java", "java.nio.file.Paths" ).get( JavaCast( "string", variables.filepath ) );
			return CreateObject( "java", "java.nio.file.Files" ).newBufferedWriter( path, charset );
		}
		var path = CreateObject( "java", "java.nio.file.Paths" ).get( JavaCast( "string", variables.filepath ), [] );
		return CreateObject( "java", "java.nio.file.Files" ).newBufferedWriter( path, charset, [] );
	}

	private void function printFromArray( required printer ){
		if( useParallelThreads() ){
			var printRowFunction = function( row ){
				printRowFromArray( row, printer );//don't scope
			};
			printUsingParallelThreads( printRowFunction );
			return;
		}
		for( var row in variables.data ){
			printRowFromArray( row, arguments.printer );
		}
	}

	private void function printFromQuery( required printer ){
		var columns = variables.library.getQueryHelper()._QueryColumnArray( variables.data );
		if( useParallelThreads() ){
			var printRowFunction = function( row ){
				printRowFromQuery( row, columns, printer );//don't scope
			};
			printUsingParallelThreads( printRowFunction );
			return;
		}
		for( var row in variables.data ){
			printRowFromQuery( row, columns, arguments.printer );
		}
	}

	private void function printUsingParallelThreads( required function printRowFunction ){
		variables.data.Each(
			arguments.printRowFunction
			,true
			,variables.parallelThreadsToUse
		);
	}

	private function printRowFromArray( required row, required printer ){
		if( IsStruct( arguments.row ) )
			arguments.row = _StructValueArray( arguments.row );
		arguments.row = checkArrayRow( arguments.row );
		printRow( arguments.printer, arguments.row );
	}

	private function printRowFromQuery( required row, required columns, required printer ){
		arguments.row = convertQueryRowToArray( arguments.row, arguments.columns );
		printRow( arguments.printer, arguments.row );
	}

	private void function printRow( required printer, required array row ){
		arguments.printer.printRecord( JavaCast( "string[]", row ) );//force numbers to strings to avoid 0.0 formatting
	}

	private array function checkArrayRow( required array row ){
		var totalColumns = arguments.row.Len();
		cfloop( from=1, to=totalColumns, index="i" ){
			var value = arguments.row[ i ];
			if( !IsSimpleValue( value ) )
				Throw( type=variables.library.getExceptionType() & ".invalidDataForCsv", message="Invalid data", detail="Your data contains complex values which cannot be output to CSV" );
			arguments.row[ i ] = formatDateString( value );
		}
		return arguments.row;
	}

	private string function formatDateString( required string value ){
		if( !variables.library.getDateHelper().isDateObject( arguments.value ) )
			return arguments.value;
		return DateTimeFormat( arguments.value, variables.library.getDateFormats().DATETIME );
	}

	private array function convertQueryRowToArray( required struct row, required array columns ){
		var result = [];
		for( var column IN arguments.columns ){
			var cellValue = formatDateString( arguments.row[ column ] );
			result.Append( cellValue );
		};
		return result;
	}

	private boolean function useParallelThreads(){
		return ( variables.parallelThreadsToUse > 1 );
	}

	private array function _StructValueArray( required struct data ){
		try{
			return StructValueArray( arguments.data ); // Lucee 5.3.8.117+
		}
		catch( any exception ){
			if( !exception.message.REFindNoCase( "undefined|no matching function" ) )
				rethrow;
			var result = [];
			for( var key in arguments.data ){
				result.Append( arguments.data[ key ] );
			}
			return result;
		}
	}

}