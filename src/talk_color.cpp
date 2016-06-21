#include <Rcpp.h>
#include "color_spec.h"
using namespace Rcpp;

// [[Rcpp::export]]
std::string solid_fill(int col_font_r, int col_font_g, int col_font_b, int col_font_a) {
  color_spec cs_(col_font_r, col_font_g, col_font_b, col_font_a);
  return cs_.solid_fill();
}

// [[Rcpp::export]]
std::string color_css(int col_font_r, int col_font_g, int col_font_b, int col_font_a) {
  color_spec cs_(col_font_r, col_font_g, col_font_b, col_font_a);
  return cs_.get_css();
}
