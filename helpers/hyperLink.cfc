component extends="base" accessors="true"{

	property name="cachedDefaultStyleObjects" type="struct";

	hyperLink function init( required Spreadsheet libraryInstance ){
		variables.cachedDefaultStyleObjects = {};
		return super.init( arguments.libraryInstance );
	}

	any function throwErrorIfTypeIsInvalid( required string type ){
		var validTypes = validTypes();
		if( !validTypes.Find( arguments.type ) )
			Throw( type=library().getExceptionType() & ".invalidTypeArgument", message="Invalid type argument: '#arguments.type#'", detail="The type must be one of the following: #validTypes.ToList( ', ' )#." );
		return this;
	}

	any function throwErrorIfTooltipAndWorkbookIsXls( required workbook ){
		if( arguments.KeyExists( "tooltip" ) && !library().isXmlFormat( arguments.workbook ) )
			Throw( type=library().getExceptionType() & ".invalidSpreadsheetType", message="Invalid spreadsheet type", detail="Hyperlink tooltips can only be added to XLSX spreadsheets." );
		return this;
	}

	any function addHyperLinkToCell( required cell, required workbook, required string link, required string type, string tooltip ){
		var hyperlinkType = getClassHelper().loadClass( "org.apache.poi.common.usermodel.HyperlinkType" );
		var hyperLink = arguments.workbook.getCreationHelper().createHyperlink( hyperlinkType[ arguments.type ] );
		hyperLink.setAddress( JavaCast( "string", arguments.link ) );
		if( arguments.KeyExists( "tooltip" ) )
			hyperLink.setTooltip( JavaCast( "string", arguments.tooltip ) );
		arguments.cell.setHyperlink( hyperLink );
		return this;
	}

	any function defaultCellStyle( required workbook ){
		var spreadsheetType = library().isXmlFormat( arguments.workbook )? "xlsx": "xls";
		if( !variables.cachedDefaultStyleObjects.KeyExists( spreadsheetType ) )
			cacheDefaultCellStyle( arguments.workbook, spreadsheetType );
		return variables.cachedDefaultStyleObjects[ spreadsheetType ];
	}

 	void function setHyperLinkStyle( required workbook, required cell ){
		var defaultHyperLinkStyle = defaultCellStyle( arguments.workbook );
		try{
			arguments.cell.setCellStyle( defaultHyperLinkStyle );
		}
		catch( any exception ){
			if( exception.message CONTAINS "Style does not belong to the supplied Workbook" ){
				var newDefaultHyperLinkCellStyleForThisWorkbook = arguments.workbook.createCellStyle();
				newDefaultHyperLinkCellStyleForThisWorkbook.cloneStyleFrom( defaultHyperLinkStyle );
				arguments.cell.setCellStyle( newDefaultHyperLinkCellStyleForThisWorkbook );
				// cache for future use
				var spreadsheetType = library().isXmlFormat( arguments.workbook )? "xlsx": "xls";
				variables.cachedDefaultStyleObjects[ spreadsheetType ] = newDefaultHyperLinkCellStyleForThisWorkbook;
			}
			else
				rethrow;
		}
	}

	/* PRIVATE */

	private array function validTypes(){
		return [ "URL", "EMAIL", "FILE", "DOCUMENT" ];
	}

	private void function cacheDefaultCellStyle( required workbook, required string spreadsheetType ){
		variables.cachedDefaultStyleObjects[ arguments.spreadsheetType ] = getFormatHelper().buildCellStyle( arguments.workbook, { color: "0000ff", underline: true } );
	}

}