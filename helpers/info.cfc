component extends="base"{

	void function addInfoBinary( required workbook, required struct info ){
		arguments.workbook.createInformationProperties(); // creates the following if missing
		var documentSummaryInfo = arguments.workbook.getDocumentSummaryInformation();
		var summaryInfo = arguments.workbook.getSummaryInformation();
		for( var key in arguments.info )
			addInfoItemBinary( arguments.info, key, summaryInfo, documentSummaryInfo );
	}

	void function addInfoXml( required workbook, required struct info ){
		var workbookProperties = library().isStreamingXmlFormat( arguments.workbook )? arguments.workbook.getXSSFWorkbook().getProperties(): arguments.workbook.getProperties();
		var documentProperties = workbookProperties.getExtendedProperties().getUnderlyingProperties();
		var coreProperties = workbookProperties.getCoreProperties();
		for( var key in arguments.info )
			addInfoItemXml( arguments.info, key, documentProperties, coreProperties );
	}
	
	struct function binaryInfo( required workbook ){
		var documentProperties = arguments.workbook.getDocumentSummaryInformation();
		var coreProperties = arguments.workbook.getSummaryInformation();
		return {
			author: coreProperties.getAuthor()?:""
			,category: documentProperties.getCategory()?:""
			,comments: coreProperties.getComments()?:""
			,creationDate: coreProperties.getCreateDateTime()?:""
			,lastEdited: ( coreProperties.getEditTime() == 0 )? "": CreateObject( "java", "java.util.Date" ).init( coreProperties.getEditTime() )
			,subject: coreProperties.getSubject()?:""
			,title: coreProperties.getTitle()?:""
			,lastAuthor: coreProperties.getLastAuthor()?:""
			,keywords: coreProperties.getKeywords()?:""
			,lastSaved: coreProperties.getLastSaveDateTime()?:""
			,manager: documentProperties.getManager()?:""
			,company: documentProperties.getCompany()?:""
		};
	}

	struct function xmlInfo( required workbook ){
		var workbookProperties = library().isStreamingXmlFormat( arguments.workbook )? arguments.workbook.getXSSFWorkbook().getProperties(): arguments.workbook.getProperties();
		var documentProperties = workbookProperties.getExtendedProperties().getUnderlyingProperties();
		var coreProperties = workbookProperties.getCoreProperties();
		var result = {
			author: coreProperties.getCreator()?:""
			,category: coreProperties.getCategory()?:""
			,comments: coreProperties.getDescription()?:""
			,creationDate: coreProperties.getCreated()?:""
			,lastEdited: coreProperties.getModified()?:""
			,subject: coreProperties.getSubject()?:""
			,title: coreProperties.getTitle()?:""
			,keywords: coreProperties.getKeywords()?:""
			,lastSaved: ""// not available in xml
			,manager: documentProperties.getManager()?:""
			,company: documentProperties.getCompany()?:""
		};
		// lastAuthor is a java.util.Option object with different behaviour
		if( coreProperties.getUnderlyingProperties().getLastModifiedByProperty().isPresent() )
			result.lastAuthor = coreProperties.getUnderlyingProperties().getLastModifiedByProperty().get();
		return result;
	}

	/* Private */

	private void function addInfoItemBinary(
		required struct info
		,required string key
		,required summaryInfo
		,required documentSummaryInfo
	){
		var value = JavaCast( "string", arguments.info[ arguments.key ] );
		switch( arguments.key ){
			case "author": arguments.summaryInfo.setAuthor( value );
				return;
			case "category": arguments.documentSummaryInfo.setCategory( value );
				return;
			case "lastauthor": arguments.summaryInfo.setLastAuthor( value );
				return;
			case "comments": arguments.summaryInfo.setComments( value );
				return;
			case "keywords": arguments.summaryInfo.setKeywords( value );
				return;
			case "manager": arguments.documentSummaryInfo.setManager( value );
				return;
			case "company": arguments.documentSummaryInfo.setCompany( value );
				return;
			case "subject": arguments.summaryInfo.setSubject( value );
				return;
			case "title": arguments.summaryInfo.setTitle( value );
		}
	}

	private void function addInfoItemXml(
		required struct info
		,required string key
		,required documentProperties
		,required coreProperties
	){
		var value = JavaCast( "string", arguments.info[ key ] );
		switch( arguments.key ){
			case "author": arguments.coreProperties.setCreator( value  );
				return;
			case "category": arguments.coreProperties.setCategory( value );
				return;
			case "lastauthor": arguments.coreProperties.getUnderlyingProperties().setLastModifiedByProperty( value );
				return;
			case "comments": arguments.coreProperties.setDescription( value );
				return;
			case "keywords": arguments.coreProperties.setKeywords( value );
				return;
			case "subject": arguments.coreProperties.setSubjectProperty( value );
				return;
			case "title": arguments.coreProperties.setTitle( value );
				return;
			case "manager": arguments.documentProperties.setManager( value );
				return;
			case "company": arguments.documentProperties.setCompany( value );
		}
	}

}