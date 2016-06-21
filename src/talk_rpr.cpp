#include <Rcpp.h>
#include "rpr.h"
#include "ppr.h"
using namespace Rcpp;

// [[Rcpp::export]]
std::string w_rpr(double size, bool italic, bool bold, bool underlined,
                  int col_font_r, int col_font_g, int col_font_b, int col_font_a,
                  int col_shading_r, int col_shading_g, int col_shading_b, int col_shading_a,
                  std::string fontname, std::string vertical_align) {

  rpr rpr_(size, italic, bold, underlined,
      col_font_r, col_font_g, col_font_b, col_font_a,
      col_shading_r, col_shading_g, col_shading_b, col_shading_a,
      fontname, vertical_align );
  return rpr_.w_tag();
}

// [[Rcpp::export]]
std::string a_rpr(double size, bool italic, bool bold, bool underlined,
                  int col_font_r, int col_font_g, int col_font_b, int col_font_a,
                  int col_shading_r, int col_shading_g, int col_shading_b, int col_shading_a,
                  std::string fontname, std::string vertical_align) {

  rpr rpr_(size, italic, bold, underlined,
           col_font_r, col_font_g, col_font_b, col_font_a,
           col_shading_r, col_shading_g, col_shading_b, col_shading_a,
           fontname, vertical_align );
  return rpr_.a_tag();
}

// [[Rcpp::export]]
std::string css_rpr(double size, bool italic, bool bold, bool underlined,
                  int col_font_r, int col_font_g, int col_font_b, int col_font_a,
                  int col_shading_r, int col_shading_g, int col_shading_b, int col_shading_a,
                  std::string fontname, std::string vertical_align) {

  rpr rpr_(size, italic, bold, underlined,
           col_font_r, col_font_g, col_font_b, col_font_a,
           col_shading_r, col_shading_g, col_shading_b, col_shading_a,
           fontname, vertical_align );
  return rpr_.css();
}

