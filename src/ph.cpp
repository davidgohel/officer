#include <Rcpp.h>
#include "ph.h"
#include "color_spec.h"
#include <iostream>

using namespace Rcpp;


std::string ph::a_tag()
{
  color_spec col_(this->red, this->green, this->blue, this->alpha);

  std::stringstream os;

  os << "<p:nvSpPr><p:cNvPr id=\"0\" name=\"\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph/></p:nvPr></p:nvSpPr>";
  os << "<p:spPr>";
  os << "<a:xfrm rot=\"" << this->rot << "\"><a:off x=\"" << this->offx << "\" y=\"" << this->offy << "\"/><a:ext cx=\"" << this->cx << "\" cy=\"" << this->cy << "\"/></a:xfrm>";
  if( col_.is_transparent() ) {
    os << "<a:noFill/>";
  } else os << col_.solid_fill();
  os << "</p:spPr>";


  return os.str();
}


ph::ph(int offx, int offy, int cx, int cy, int rot, int r, int g, int b, int a):
  offx(offx), offy(offy), cx(cx), cy(cy),
  rot(rot),
  red(r), green(g), blue(b), alpha(a){

}


// [[Rcpp::export]]
std::string p_ph(int offx, int offy, int cx, int cy, int rot, int r, int g, int b, int a ) {
  ph ph_(offx, offy, cx, cy, rot, r, g, b, a);
  return ph_.a_tag();
}
