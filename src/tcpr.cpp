#include <Rcpp.h>
#include "tcpr.h"
#include <iostream>
#include "color_spec.h"
#include <gdtools.h>

using namespace Rcpp;

std::string tcpr::w_tag()
{
  color_spec shading_(this->shading_r, this->shading_g, this->shading_b, this->shading_a);

  std::stringstream os;
  os << "<w:tcPr>";

  os << "<w:tcBorders>";
  os << b.w_tag("bottom");
  os << t.w_tag("top");
  os << l.w_tag("left");
  os << r.w_tag("right");
  os << "</w:tcBorders>";
  if( !shading_.is_transparent() ) {
    os << shading_.w_shd();
  }

  os << "<w:tcMar>";
  os << "<w:top w:w=\""<< mt*20 << "\" w:type=\"dxa\"/>";
  os << "<w:bottom w:w=\""<< mb*20 << "\" w:type=\"dxa\"/>";
  os << "<w:left w:w=\""<< ml*20 << "\" w:type=\"dxa\"/>";
  os << "<w:right w:w=\""<< mr*20 << "\" w:type=\"dxa\"/>";
  os << "</w:tcMar>";

  if( text_direction == "btlr")
    os << "<w:textDirection w:val=\"btLr\"/>";
  else if( text_direction == "tbrl")
    os << "<w:textDirection w:val=\"tbRl\"/>";

  os << "<w:vAlign w:val=\"";
  if( vertical_align == "center" )
    os << "center";
  else if( vertical_align == "top" )
    os << "top";
  else os << "bottom";
  os << "\"/>";

  os << "</w:tcPr>";

  return os.str();
}




std::string tcpr::css()
{
  color_spec shading_(this->shading_r, this->shading_g, this->shading_b, this->shading_a);

  std::stringstream os;

  os << b.css("bottom");
  os << t.css("top");
  os << l.css("left");
  os << r.css("right");
  if( shading_.is_visible() > 0 )
    os << "background-color:" << shading_.get_css() << ";";
  else os << "background-color:transparent;";
  if( this->do_bgimg ){
    std::string base64_str = gdtools::base64_file_encode(this->bgimg_path);
    os << "background-image:url(data:image/png;base64," << base64_str << ");";
  }

  os << "margin-top:" << mt << "pt;";
  os << "margin-bottom:" << mb << "pt;";
  os << "margin-left:" << ml << "pt;";
  os << "margin-right:" << mr << "pt;";

  os << "vertical-align:";
  if( vertical_align == "center" )
    os << "middle;";
  else if( vertical_align == "top" )
    os << "top;";
  else os << "bottom;";

  return os.str();
}

std::string tcpr::a_tag()
{

  color_spec shading_(this->shading_r, this->shading_g, this->shading_b, this->shading_a);

  std::stringstream os;
  os << "<a:tcPr ";

  if( text_direction == "btlr")
    os << "vert=\"vert270\" ";
  else if( text_direction == "tbrl")
    os << "vert=\"vert\" ";

  os << "anchor=\"";
  if( vertical_align == "center" )
    os << "ctr";
  else if( vertical_align == "top" )
    os << "t";
  else
    os << "b";
  os << "\" ";

  os << "marB=\""<< mb*12700 << "\" ";
  os << "marT=\""<< mt*12700 << "\" ";
  os << "marR=\""<< mr*12700 << "\" ";
  os << "marL=\""<< ml*12700 << "\">";

  os << l.a_tag("L");
  os << r.a_tag("R");
  os << t.a_tag("T");
  os << b.a_tag("B");

  if( !do_bgimg && !shading_.is_transparent() ) {
    os << shading_.solid_fill();
  } else if( do_bgimg ) {
    os << "<a:blipFill rotWithShape=\"1\"><a:blip r:embed=\"";
    os << bgimg_rid;
    os << "\"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>";
  }

  os << "</a:tcPr>";

  return os.str();
}
tcpr::tcpr(std::string vertical_align, std::string text_direction,
         int mb, int mt, int ml, int mr,
         int shd_r, int shd_g, int shd_b, int shd_a,
         bool do_bgimg, std::string bgimg_rid, std::string bgimg_path,
         IntegerVector btlr_red, IntegerVector btlr_green,
         IntegerVector btlr_blue, IntegerVector btlr_alpha,
         CharacterVector type, IntegerVector width):
  vertical_align(vertical_align), text_direction(text_direction),
  mb(mb), mt(mt), ml(ml), mr(mr),
  shading_r(shd_r), shading_g(shd_g), shading_b(shd_b), shading_a(shd_a),
  do_bgimg(do_bgimg), bgimg_rid(bgimg_rid), bgimg_path(bgimg_path),
  b(btlr_red[0], btlr_green[0], btlr_blue[0], btlr_alpha[0], as<std::string>(type[0]), width[0]),
  t(btlr_red[1], btlr_green[1], btlr_blue[1], btlr_alpha[1], as<std::string>(type[1]), width[1]),
  l(btlr_red[2], btlr_green[2], btlr_blue[2], btlr_alpha[2], as<std::string>(type[2]), width[2]),
  r(btlr_red[3], btlr_green[3], btlr_blue[3], btlr_alpha[3], as<std::string>(type[3]), width[3]){
}
