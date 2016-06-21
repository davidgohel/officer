#include <Rcpp.h>
#include "color_spec.h"
#include <iostream>
#include "R_ext/GraphicsDevice.h"
using namespace Rcpp;

std::string color_spec::solid_fill()
{

  char col_buf[ 100 ];
  sprintf( col_buf, "%02X%02X%02X", this->red, this->green, this->blue);
  std::string col_str = col_buf;

  std::stringstream os;
  os << "<a:solidFill><a:srgbClr val=\"";
  os << col_str;
  os << "\">";
  os << "<a:alpha val=\"" << (int)(this->alpha / 255.0 * 100000) << "\"/>";
  os << "</a:srgbClr></a:solidFill>";
  return os.str();
}

std::string color_spec::solid_fill_w14()
{

  char col_buf[ 100 ];
  sprintf( col_buf, "%02X%02X%02X", this->red, this->green, this->blue);
  std::string col_str = col_buf;

  std::stringstream os;
  os << "<w14:solidFill><w14:srgbClr val=\"";
  os << col_str;
  os << "\">";
  os << "<w14:alpha val=\"" << (int)(this->alpha / 255.0 * 100000) << "\"/>";
  os << "</w14:srgbClr></w14:solidFill>";
  return os.str();
}
std::string color_spec::w_color()
{

  char col_buf[ 100 ];
  sprintf( col_buf, "%02X%02X%02X", this->red, this->green, this->blue);
  std::string col_str = col_buf;

  std::stringstream os;
  os << "<w:color w:val=\"";
  os << col_str;
  os << "\"/>";

  return os.str();
}

std::string color_spec::w_shd()
{
  char col_buf[ 100 ];
  sprintf( col_buf, "%02X%02X%02X", this->red, this->green, this->blue);
  std::string col_str = col_buf;
  std::stringstream os;
  os << "<w:shd w:fill=\"";
  os << col_str;
  os << "\"/>";

  return os.str();
}

std::string color_spec::get_hex()
{
  char col_buf[ 100 ];
  sprintf( col_buf, "%02X%02X%02X", this->red, this->green, this->blue);
  std::string col_str = col_buf;
  std::stringstream os;
  os << col_str;

  return os.str();
}

std::string color_spec::get_css()
{
  char col_buf[ 100 ];
  sprintf( col_buf, "rgba(%d,%d,%d,%.2f)", this->red, this->green, this->blue, (double)this->alpha/255);
  std::string col_str = col_buf;
  std::stringstream os;
  os << col_str;

  return os.str();
}
int color_spec::is_visible() {
  return (this->alpha != 0);
}
int color_spec::has_alpha() {
  return (this->alpha < 255);
}

int color_spec::is_transparent() {
  return (this->alpha == 0);
}

color_spec::color_spec (int r, int g, int b, int a ):
  red(r), green(g), blue(b), alpha(a){
}
