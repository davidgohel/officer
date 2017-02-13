#include <Rcpp.h>
using namespace Rcpp;

// [[Rcpp::export]]
std::string wml_run_pic(std::string src, double width, double height) {
  std::stringstream os;
  os << "<w:r><w:rPr/><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">";
  os << "<wp:extent cx=\"" << (int)(width * 12700) << "\" cy=\"" << (int)(height * 12700) << "\"/>";
  os << "<wp:docPr id=\"\" name=\"\"/>";
  os << "<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/></wp:cNvGraphicFramePr>";
  os << "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">";
  os << "<pic:nvPicPr>";
  os << "<pic:cNvPr id=\"\" name=\"\"/>";
  os << "<pic:cNvPicPr><a:picLocks noChangeAspect=\"1\" noChangeArrowheads=\"1\"/>";
  os << "</pic:cNvPicPr></pic:nvPicPr>";
  os << "<pic:blipFill>";
  os << "<a:blip r:embed=\"" << src << "\"/>";
  os << "<a:srcRect/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>";
  os << "<pic:spPr bwMode=\"auto\"><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"" << (int)(width * 12700) << "\" cy=\"" << (int)(height * 12700) << "\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></pic:spPr>";
  os << "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";
    return os.str();
}

// [[Rcpp::export]]
std::string pml_run_pic(std::string src, double width, double height) {
  std::stringstream os;

  os << "<p:pic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">";
  os << "<p:nvPicPr>";
  os << "<p:cNvPr id=\"\" name=\"pic\"/>";
  os << "<p:cNvPicPr/>";
  os << "<p:nvPr/>";
  os << "</p:nvPicPr>";
  os << "<p:blipFill>";
  os << "<a:blip cstate=\"print\" r:embed=\"" << src << "\"/>";
  os << "<a:stretch><a:fillRect/></a:stretch>";
  os << "</p:blipFill>";

  os << "<p:spPr>";

  os << "<a:xfrm><a:off x=\"\" y=\"\"/><a:ext cx=\"" << (int)(width * 12700) << "\" cy=\"" << (int)(height * 12700) << "\"/></a:xfrm>";
  os << "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>";
  os << "</p:spPr>";
  os << "</p:pic>";

  return os.str();
}

