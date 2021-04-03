/**
* Adapted from https://stackoverflow.com/questions/51077404/apache-poi-adding-watermark-in-excel-workbook/51103756#51103756
**/
package luceeSpreadsheet;

import java.io.*;
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.ooxml.*;
import org.apache.xmlbeans.*;

import static org.apache.poi.ooxml.POIXMLTypeLoader.DEFAULT_XML_OPTIONS;

public class HeaderImageVML extends POIXMLDocumentPart {

  String rIdPic = "";
  java.awt.Dimension imageDimension = null;
  String position = "";

  public HeaderImageVML(PackagePart part) {
    super(part);
  }

  public HeaderImageVML setRIdPic(String rIdPic) {
    this.rIdPic = rIdPic;
    return this;
  }

  public HeaderImageVML setPosition(String position) {
    this.position = position;
    return this;
  }

  public HeaderImageVML setImageDimension(java.awt.Dimension imageDimension) {
   this.imageDimension = imageDimension;
   return this;
  }

  @Override
  protected void commit() throws IOException {
    PackagePart part = getPackagePart();
    OutputStream out = part.getOutputStream();
    try {
      XmlObject doc = XmlObject.Factory.parse(
        "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\""
        +" xmlns:o=\"urn:schemas-microsoft-com:office:office\""
        +" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">"
        +" <o:shapelayout v:ext=\"edit\">"
        +"  <o:idmap v:ext=\"edit\" data=\"1\"/>"
        +" </o:shapelayout><v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\""
        +"  o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">"
        +"  <v:stroke joinstyle=\"miter\"/>"
        +"  <v:formulas>"
        +"   <v:f eqn=\"if lineDrawn pixelLineWidth 0\"/>"
        +"   <v:f eqn=\"sum @0 1 0\"/>"
        +"   <v:f eqn=\"sum 0 0 @1\"/>"
        +"   <v:f eqn=\"prod @2 1 2\"/>"
        +"   <v:f eqn=\"prod @3 21600 pixelWidth\"/>"
        +"   <v:f eqn=\"prod @3 21600 pixelHeight\"/>"
        +"   <v:f eqn=\"sum @0 0 1\"/>"
        +"   <v:f eqn=\"prod @6 1 2\"/>"
        +"   <v:f eqn=\"prod @7 21600 pixelWidth\"/>"
        +"   <v:f eqn=\"sum @8 21600 0\"/>"
        +"   <v:f eqn=\"prod @7 21600 pixelHeight\"/>"
        +"   <v:f eqn=\"sum @10 21600 0\"/>"
        +"  </v:formulas>"
        +"  <v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/>"
        +"  <o:lock v:ext=\"edit\" aspectratio=\"t\"/>"
        +" </v:shapetype><v:shape id=\"" + position + "\" o:spid=\"_x0000_s1025\" type=\"#_x0000_t75\""
        +"  style='position:absolute;margin-left:0;margin-top:0;"
        +"width:" + (int)imageDimension.getWidth() + "px;height:" + (int)imageDimension.getHeight() + "px;"
        +"z-index:1'>"
        +"  <v:imagedata o:relid=\""+ rIdPic + "\"/>"
        +"  <o:lock v:ext=\"edit\" rotation=\"t\"/>"
        +" </v:shape></xml>"
      );
      doc.save(out, DEFAULT_XML_OPTIONS);
      out.close();
    } catch (Exception ex) {
      ex.printStackTrace();
    }
  }

 }