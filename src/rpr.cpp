#include <Rcpp.h>
#include "rpr.h"
#include "color_spec.h"
#include <iostream>

using namespace Rcpp;

std::string rpr::a_tag()
{
  color_spec col_(this->col_font_r, this->col_font_g, this->col_font_b, this->col_font_a);
  if( col_.is_visible() < 1 ) return "";

  std::stringstream os;

  os << "<a:rPr";
  if( this->size > 0 ){
    os << " sz=\"";
    os << (int)(this->size*100);
    os << "\"";
  }

  if( this->italic ) os << " i=\"1\"";
  if( this->bold ) os << " b=\"1\"";
  if( this->underlined ) os << " u=\"1\"";
  os << ">";
  os << col_.solid_fill();
  os << "<a:latin typeface=\"" << this->fontname << "\"/>";
  os << "<a:cs typeface=\"" << this->fontname << "\"/>";

  os << "</a:rPr>";

  return os.str();
}

std::string rpr::w_tag()
{
  color_spec col_(this->col_font_r, this->col_font_g, this->col_font_b, this->col_font_a);
  color_spec shading_(this->col_shading_r, this->col_shading_g, this->col_shading_b, this->col_shading_a);
  if( col_.is_visible() < 1 ) return "";

  std::stringstream os;

  os << "<w:rPr>";
  os << "<w:rFonts";
  os << " w:ascii=\"" << this->fontname << "\"";
  os << " w:hAnsi=\"" << this->fontname << "\"";
  os << " w:cs=\"" << this->fontname << "\"";
  os << "/>";

  if( this->italic ) os << "<w:i/>";
  if( this->bold ) os << "<w:b/>";
  if( this->underlined ) os << "<w:u/>";

  os << "<w:sz w:val=\"";
  os << (int)(this->size*2);
  os << "\"/>";
  os << "<w:szCs w:val=\"";
  os << (int)(this->size*2);
  os << "\"/>";

  os << col_.w_color();

  if( shading_.is_visible() > 0 )
    os << shading_.w_shd();

  if (col_.has_alpha() > 0) {
    os << "<w14:textFill>";
    os << col_.solid_fill_w14();
    os << "</w14:textFill>";
  }
  os << "</w:rPr>";
  return os.str();
}


std::string rpr::css()
{
  color_spec col_(this->col_font_r, this->col_font_g, this->col_font_b, this->col_font_a);
  color_spec shading_(this->col_shading_r, this->col_shading_g, this->col_shading_b, this->col_shading_a);
  if( col_.is_visible() < 1 ) return "";

  std::stringstream os;

  os << "font-family:'" << this->fontname << "';";
  os << "color:" << col_.get_css() << ";";
  os << "font-size:" << (int)this->size << "px;";

  if( this->italic ) os << "font-style:italic;";
  else os << "font-style:normal;";
  if( this->bold ) os << "font-weight:bold;";
  else os << "font-weight:normal;";

  if( this->underlined ) os << "text-decoration:underline;";
  else os << "text-decoration:none;";

  if( shading_.is_visible() > 0 )
    os << "background-color:" << shading_.get_css() << ";";
  else os << "background-color:transparent;";

  return os.str();
}

rpr::rpr(double size, bool italic, bool bold, bool underlined,
          int col_font_r, int col_font_g, int col_font_b, int col_font_a,
          int col_shading_r, int col_shading_g, int col_shading_b, int col_shading_a,
          std::string fontname, std::string vertical_align ):
  size(size),
  italic(italic),
  bold(bold),
  underlined(underlined),
  vertical_align(vertical_align),
  col_font_r(col_font_r),
  col_font_g(col_font_g),
  col_font_b(col_font_b),
  col_font_a(col_font_a),
  col_shading_r(col_shading_r),
  col_shading_g(col_shading_g),
  col_shading_b(col_shading_b),
  col_shading_a(col_shading_a),
  fontname(fontname){
}



