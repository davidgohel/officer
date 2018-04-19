#include <Rcpp.h>
#include "border.h"
#include "color_spec.h"
#include <iostream>
#include <iomanip>

using namespace Rcpp;

std::string border::w_tag(std::string side)
{
  color_spec col_(this->red, this->green, this->blue, this->alpha);

  if( col_.is_transparent() || width < .001 || type == "none") {
    return "";
  }

  std::stringstream os;
  os << "<w:" << side << " ";

  os << "w:val=";
  if( type == "solid")
    os << "\"single\" ";
  else if( type == "dotted")
    os << "\"dotted\" ";
  else if( type == "dashed")
    os << "\"dashed\" ";

  os << "w:sz=\"";
  os << std::setprecision(0) << std::fixed << width * 8;
  os << "\" ";
  os << "w:space=\"0\" ";

  os << "w:color=\"" << col_.get_hex() << "\" ";

  os << "/>";

  return os.str();
}

std::string border::a_tag(std::string side)
{
  color_spec col_(this->red, this->green, this->blue, this->alpha);

  /*if( col_.is_transparent() || width < 1 ) {
    return "";
  }*/

  std::stringstream os;

  os << "<a:ln" << side << " algn=\"ctr\" cmpd=\"sng\" cap=\"flat\" ";
  os << "w=\"";
  os << std::setprecision(0) << std::fixed << width * 12700;
  os << "\">";

  if( col_.is_transparent() || width < .001 ) {
    os << "<a:noFill/>";
  } else os << col_.solid_fill();

  os << "<a:prstDash val=";
  if( type == "solid")
    os << "\"solid\"/>";
  else if( type == "dotted")
    os << "\"sysDot\"/>";
  else if( type == "dashed")
    os << "\"sysDash\"/>";

  os << "</a:ln" << side << ">";
  return os.str();
}


std::string border::css(std::string side)
{
  color_spec col_(this->red, this->green, this->blue, this->alpha);
  std::stringstream os;

  os << "border-" << side << ": ";
  os << std::setprecision(2) << std::fixed << width << "pt ";

  if( type == "dotted")
    os << "dotted ";
  else if( type == "dashed")
    os << "dashed ";
  else os << "solid ";

  if( col_.is_transparent() ) {
    os << "transparent;";
  } else os << col_.get_css() << ";";

  return os.str();
}

border::border(int r, int g, int b, int a,
               std::string type, double width):
  red(r), green(g), blue(b), alpha(a),
  type(type), width(width){

}
