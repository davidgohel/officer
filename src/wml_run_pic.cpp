#include <Rcpp.h>
using namespace Rcpp;

// [[Rcpp::export]]
std::string wml_run_pic(std::string src, double width, double height) {
  std::stringstream os;
  os << "<w:r><w:rPr/><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">";
  os << "<wp:extent cx=\"" << (int)(width * 12700) << "\" cy=\"" << (int)(height * 12700) << "\"/>";
  os << "<wp:docPr id=\"DRAWINGOBJECTID\" name=\"\"/>";
  os << "<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/></wp:cNvGraphicFramePr>";
  os << "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">";
  os << "<pic:nvPicPr>";
  os << "<pic:cNvPr id=\"PICTUREID\" name=\"\"/>";
  os << "<pic:cNvPicPr><a:picLocks noChangeAspect=\"1\" noChangeArrowheads=\"1\"/>";
  os << "</pic:cNvPicPr></pic:nvPicPr>";
  os << "<pic:blipFill>";
  os << "<a:blip r:embed=\"" << src << "\"/>";
  os << "<a:srcRect/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>";
  os << "<pic:spPr bwMode=\"auto\"><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"" << (int)(width * 12700) << "\" cy=\"" << (int)(height * 12700) << "\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></pic:spPr>";
  os << "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";
    return os.str();
}

