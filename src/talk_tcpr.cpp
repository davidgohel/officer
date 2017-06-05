#include <Rcpp.h>
#include "tcpr.h"
using namespace Rcpp;


// [[Rcpp::export]]
std::string w_tcpr(std::string vertical_align, std::string text_direction,
                   int mb, int mt, int ml, int mr,
                   int shd_r, int shd_g, int shd_b, int shd_a,
                   bool do_bgimg, std::string bgimg_rid, std::string bgimg_path,
                   IntegerVector btlr_red, IntegerVector btlr_green,
                   IntegerVector btlr_blue, IntegerVector btlr_alpha,
                   CharacterVector type, IntegerVector width ) {
  tcpr tcpr_(vertical_align, text_direction,
             mb, mt, ml, mr,
             shd_r, shd_g, shd_b, shd_a,
             do_bgimg, bgimg_rid, bgimg_path,
             btlr_red, btlr_green, btlr_blue, btlr_alpha,
             type, width);
  return tcpr_.w_tag();
}

// [[Rcpp::export]]
std::string a_tcpr(std::string vertical_align,std::string text_direction,
                   int mb, int mt, int ml, int mr,
                   int shd_r, int shd_g, int shd_b, int shd_a,
                   bool do_bgimg, std::string bgimg_rid, std::string bgimg_path,
                   IntegerVector btlr_red, IntegerVector btlr_green,
                   IntegerVector btlr_blue, IntegerVector btlr_alpha,
                   CharacterVector type, IntegerVector width) {
  tcpr tcpr_(vertical_align, text_direction,
             mb, mt, ml, mr,
             shd_r, shd_g, shd_b, shd_a,
             do_bgimg, bgimg_rid, bgimg_path,
             btlr_red, btlr_green, btlr_blue, btlr_alpha,
             type, width);
  return tcpr_.a_tag();
}

// [[Rcpp::export]]
std::string css_tcpr(std::string vertical_align,std::string text_direction,
                   int mb, int mt, int ml, int mr,
                   int shd_r, int shd_g, int shd_b, int shd_a,
                   bool do_bgimg, std::string bgimg_rid, std::string bgimg_path,
                   IntegerVector btlr_red, IntegerVector btlr_green,
                   IntegerVector btlr_blue, IntegerVector btlr_alpha,
                   CharacterVector type, IntegerVector width ) {
  tcpr tcpr_(vertical_align, text_direction,
             mb, mt, ml, mr,
             shd_r, shd_g, shd_b, shd_a,
             do_bgimg, bgimg_rid, bgimg_path,
             btlr_red, btlr_green, btlr_blue, btlr_alpha,
             type, width);
  return tcpr_.css();
}


