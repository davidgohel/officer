#include <Rcpp.h>
#include "ppr.h"
#include <iostream>
#include "color_spec.h"

using namespace Rcpp;

std::string ppr::a_tag()
{
  std::stringstream os;
  os << "<a:pPr";

  if (this->text_align == "left" )
    os << " algn=\"l\"";
  else if (this->text_align == "center" )
    os << " algn=\"ctr\"";
  else
    os << " algn=\"r\"";
  os << " marL=\""<< pl*12700 << "\" marR=\""<< pr*12700 << "\">";
  os << "<a:spcBef><a:spcPts val=\""<< pt*100 << "\" /></a:spcBef>";
  os << "<a:spcAft><a:spcPts val=\""<< pb*100 << "\" /></a:spcAft>";
  os << "<a:buNone/>";
  os << "</a:pPr>";
  return os.str();
}


std::string ppr::css()
{
  color_spec shading_(this->shading_r, this->shading_g, this->shading_b, this->shading_a);

  std::stringstream os;
  os << "margin:0pt;";
  os << "text-align:" << this->text_align << ";";

  os << b.css("bottom");
  os << t.css("top");
  os << l.css("left");
  os << r.css("right");

  os << "padding-top:" << pt << "pt;";
  os << "padding-bottom:" << pb << "pt;";
  os << "padding-left:" << pl << "pt;";
  os << "padding-right:" << pr << "pt;";

  if( shading_.is_visible() > 0 )
    os << "background-color:" << shading_.get_css() << ";";
  else os << "background-color:transparent;";

  return os.str();
}

std::string ppr::w_tag()
{
  color_spec shading_(this->shading_r, this->shading_g, this->shading_b, this->shading_a);

  std::stringstream os;
  os << "<w:pPr>";
  os << "<w:jc w:val=\"" << text_align << "\"/>";
  os << b.w_tag("bottom");
  os << t.w_tag("top");
  os << l.w_tag("left");
  os << r.w_tag("right");

  os << "<w:spacing " << "w:after=\""<< pb*20 << "\" w:before=\""<< pt*20 << "\"/>";
  os << "<w:ind w:left=\""<< pl*20 << "\" w:right=\""<< pr*20 << "\"/>";

  if( !shading_.is_transparent() ) {
    os << shading_.w_shd();
  }

  os << "</w:pPr>";

  return os.str();
}

ppr::ppr(std::string text_align,
          int pb, int pt, int pl, int pr,
          int shd_r, int shd_g, int shd_b, int shd_a,
          IntegerVector btlr_red, IntegerVector btlr_green,
          IntegerVector btlr_blue, IntegerVector btlr_alpha,
          CharacterVector type, IntegerVector width):
  text_align(text_align),
  pb(pb), pt(pt), pl(pl), pr(pr),
  shading_r(shd_r), shading_g(shd_g), shading_b(shd_b), shading_a(shd_a),
  b(btlr_red[0], btlr_green[0], btlr_blue[0], btlr_alpha[0], as<std::string>(type[0]), width[0]),
  t(btlr_red[1], btlr_green[1], btlr_blue[1], btlr_alpha[1], as<std::string>(type[1]), width[1]),
  l(btlr_red[2], btlr_green[2], btlr_blue[2], btlr_alpha[2], as<std::string>(type[2]), width[2]),
  r(btlr_red[3], btlr_green[3], btlr_blue[3], btlr_alpha[3], as<std::string>(type[3]), width[3]){
}
