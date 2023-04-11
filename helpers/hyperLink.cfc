component extends="base"{

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
		var hyperlinkType = library().createJavaObject( "org.apache.poi.common.usermodel.HyperlinkType" );
		var hyperLink = arguments.workbook.getCreationHelper().createHyperlink( hyperlinkType[ arguments.type ] );
		hyperLink.setAddress( JavaCast( "string", arguments.link ) );
		if( arguments.KeyExists( "tooltip" ) )
			hyperLink.setTooltip( JavaCast( "string", arguments.tooltip ) );
		arguments.cell.setHyperlink( hyperLink );
		return this;
	}

	any function defaultCellStyle( required workbook ){
		return getFormatHelper().getCachedCellStyle( arguments.workbook, { color: "0000ff", underline: true } );
	}

 	void function setHyperLinkDefaultStyle( required workbook, required cell ){
		getFormatHelper().setCellStyle( arguments.cell, defaultCellStyle( arguments.workbook ) );
	}

	/* PRIVATE */

	private array function validTypes(){
		return [ "URL", "EMAIL", "FILE", "DOCUMENT" ];
	}

}