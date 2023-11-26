component extends="base"{

	//see https://stackoverflow.com/questions/51077404/apache-poi-adding-watermark-in-excel-workbook/51103756#51103756
	any function setHeaderOrFooterImage(
		required workbook
		,required string position // left|center|right
		,required any image
		,string imageType
		,boolean isHeader=true //false = footer
	){
		var headerType = arguments.isHeader? "Header": "Footer";
		if( !library().isXmlFormat( arguments.workbook ) )
			Throw( type=library().getExceptionType() & ".invalidSpreadsheetType", message="Invalid spreadsheet type", detail="#headerType# images can only be added to XLSX spreadsheets." );
		var imageIndex = getImageHelper().addImageToWorkbook( argumentCollection=arguments );
		var sheet = getSheetHelper().getActiveSheet( arguments.workbook );
		var headerObject = arguments.isHeader? sheet.getHeader(): sheet.getFooter();
		var headerTypeInitialLetter = headerType.Left( 1 ); // "H" or "F"
		var headerImagePartName = "/xl/drawings/vmlDrawing1.vml";
		switch( arguments.position ){
			case "left": case "l":
				headerObject.setLeft( "&G" ); //&G means Graphic
				var vmlPosition = "L#headerTypeInitialLetter#";
				break;
			case "center": case "c": case "centre":
				headerObject.setCenter( "&G" );
				var vmlPosition = "C#headerTypeInitialLetter#";
				break;
			case "right": case "r":
				headerObject.setRight( "&G" );
				var vmlPosition = "R#headerTypeInitialLetter#";
				break;
			default: Throw( type=library().getExceptionType() & ".invalidPositionArgument", message="Invalid #headerType.LCase()# position", detail="The 'position' argument '#arguments.position#' is invalid. Use 'left', 'center' or 'right'" );
		}
		// check for existing header/footer images
		var existingRelation = getExistingHeaderFooterImageRelation( sheet, headerImagePartName );
		var sheetHasExistingHeaderFooterImages = !IsNull( existingRelation );
		if( sheetHasExistingHeaderFooterImages ){
			var part = existingRelation.getPackagePart();
			try{
				var headerImageXML = existingRelation.getXml();//Works OK if workbook not previously saved with header/footer images
			}
			catch( any exception ){
				if( exception.message.Find( "getXml" ) )
					// ...but won't work if file has been previously saved with a header/footer image
					Throw( type=library().getExceptionType() & ".existingHeaderOrFooter", message="Spreadsheet contains an existing header or footer", detail="Header/footer images can't currently be added to spreadsheets read from disk that already have them." );
					/*
						TODO why won't this work? This is how to get the existing xml, but it won't save back modified to the vmlDrawing1.vml part for some reason
						headerImageXML = sheet.getRelations()[ i ].getDocument().xmlText();
					*/
				else
					rethrow;
			}
		}
		else{
			var OPCPackage = workbook.getPackage();
			var partName = library().createJavaObject( "org.apache.poi.openxml4j.opc.PackagingURIHelper" ).createPartName( headerImagePartName );
			var part = OPCPackage.createPart( partName, "application/vnd.openxmlformats-officedocument.vmlDrawing" );
			var headerImageXML = getNewHeaderImageXML();
		}
		var headerImageVml = library().createJavaObject( "spreadsheetCFML.HeaderImageVML" ).init( part );
		//create the relation to the picture
		var pictureData = arguments.workbook.getAllPictures().get( imageIndex );
		var xssfImageRelation = library().createJavaObject( "org.apache.poi.xssf.usermodel.XSSFRelation" ).IMAGES;
		var pictureRelationID = headerImageVml.addRelation( JavaCast( "null", 0 ), xssfImageRelation, pictureData ).getRelationship().getId();
		//get image dimension
		try{
			var imageInputStream = CreateObject( "java", "java.io.ByteArrayInputStream" ).init( pictureData.getData() );
			var imageUtils = library().createJavaObject( "org.apache.poi.ss.util.ImageUtils" );
			var imageDimension = imageUtils.getImageDimension( imageInputStream, pictureData.getPictureType() );
		}
		catch( any exception ){
			rethrow;
		}
		finally{
			getFileHelper().closeLocalFileOrStream( local, "imageInputStream" );
		}
		var newShapeElement = createNewHeaderImageVMLShape( pictureRelationID, vmlPosition, imageDimension );
		headerImageXML = headerImageXML.ReplaceAll( "(?i)(<\/[\w:]*xml>)", newShapeElement & "$1" );//Use java regex for group reference consistency
		headerImageVml.setXml( headerImageXML );
	  //create the sheet/vml relation
	  var xssfVmlRelation = library().createJavaObject( "org.apache.poi.xssf.usermodel.XSSFRelation" ).VML_DRAWINGS;
  	var sheetVmlRelationID = sheet.addRelation( JavaCast( "null", 0 ), xssfVmlRelation, headerImageVml ).getRelationship().getId();
  	//create the <legacyDrawingHF r:id="..."/> in /xl/worksheets/sheetN.xml
  	if( !sheetHasExistingHeaderFooterImages )
  		sheet.getCTWorksheet().addNewLegacyDrawingHF().setId( sheetVmlRelationID );
  	return this;
	}

	/* Private */

	private any function getExistingHeaderFooterImageRelation( required sheet, required string headerImagePartName ){
		var totalExistingRelations = arguments.sheet.getRelations().Len();
		if( totalExistingRelations == 0 )
			return;
		cfloop( from=1, to=totalExistingRelations, index="local.i" ){
			var relation = arguments.sheet.getRelations()[ i ];
			if( relation.getPackagePart().getPartName().getName() == arguments.headerImagePartName )
				return relation;
		}	
	}

	private string function createNewHeaderImageVMLShape( required string pictureRelationID, required string vmlPosition, required imageDimension ){
		return Trim( '
			<v:shape id="#arguments.vmlPosition#" o:spid="_x0000_s1025" type="##_x0000_t75" style="position:absolute;margin:0;width:#arguments.imageDimension.getWidth()#pt;height:#arguments.imageDimension.getHeight()#pt;">
				<v:imagedata o:relid="#pictureRelationID#" o:title="#pictureRelationID#" />
				<o:lock v:ext="edit" rotation="t" />
			</v:shape>
		' ).REReplace( ">\s+<", "><", "ALL" );
	}

	private string function getNewHeaderImageXML(){
		return '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1" /></o:shapelayout><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter" /><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0" /><v:f eqn="sum @0 1 0" /><v:f eqn="sum 0 0 @1" /><v:f eqn="prod @2 1 2" /><v:f eqn="prod @3 21600 pixelWidth" /><v:f eqn="prod @3 21600 pixelHeight" /><v:f eqn="sum @0 0 1" /><v:f eqn="prod @6 1 2" /><v:f eqn="prod @7 21600 pixelWidth" /><v:f eqn="sum @8 21600 0" /><v:f eqn="prod @7 21600 pixelHeight" /><v:f eqn="sum @10 21600 0" /></v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect" /><o:lock v:ext="edit" aspectratio="t" /></v:shapetype></xml>';
	}

}