/**
* Adapted from https://stackoverflow.com/questions/51077404/apache-poi-adding-watermark-in-excel-workbook/51103756#51103756
**/
package spreadsheetCFML;

import java.io.*;
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.ooxml.*;
import org.apache.xmlbeans.*;

import static org.apache.poi.ooxml.POIXMLTypeLoader.DEFAULT_XML_OPTIONS;

public class HeaderImageVML extends POIXMLDocumentPart {

  String xml = "";

  public HeaderImageVML(PackagePart part) {
    super(part);
  }

  public HeaderImageVML setXml(String xml) {
    this.xml = xml;
    return this;
  }

  public String getXml(){
    return xml;
  }

  @Override
  protected void commit() throws IOException {
    PackagePart part = getPackagePart();
    OutputStream out = part.getOutputStream();
    try {
      XmlObject doc = XmlObject.Factory.parse( xml );
      doc.save(out, DEFAULT_XML_OPTIONS);
      out.close();
    } catch (Exception ex) {
      ex.printStackTrace();
    }
  }

 }