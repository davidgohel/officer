#include <Rcpp.h>
#include "rpr.h"
#include "ppr.h"
using namespace Rcpp;

// [[Rcpp::export]]
SEXP ppr_new(Rcpp::List compounds) {
  Rcpp::XPtr<ppr> ptr( new ppr(Rcpp::as<std::string>(compounds["text_align"]),

                               Rcpp::as<int>(compounds["pb"]),
                               Rcpp::as<int>(compounds["pt"]),
                               Rcpp::as<int>(compounds["pl"]),
                               Rcpp::as<int>(compounds["pr"]),

                               Rcpp::as<int>(compounds["shd_r"]),
                               Rcpp::as<int>(compounds["shd_g"]),
                               Rcpp::as<int>(compounds["shd_b"]),
                               Rcpp::as<int>(compounds["shd_a"]),

                               Rcpp::as<IntegerVector>(compounds["btlr_red"]),
                               Rcpp::as<IntegerVector>(compounds["btlr_green"]),
                               Rcpp::as<IntegerVector>(compounds["btlr_blue"]),
                               Rcpp::as<IntegerVector>(compounds["btlr_alpha"]),
                               Rcpp::as<CharacterVector>(compounds["type"]),
                               Rcpp::as<IntegerVector>(compounds["width"])
      ), true );
  return ptr;
}

// [[Rcpp::export]]
std::string w_ppr(std::string text_align,
                  int pb, int pt, int pl, int pr,
                  int shd_r, int shd_g, int shd_b, int shd_a,
                  IntegerVector btlr_red, IntegerVector btlr_green,
                  IntegerVector btlr_blue, IntegerVector btlr_alpha,
                  CharacterVector type, IntegerVector width ) {
  ppr ppr_(text_align, pb, pt, pl, pr,
           shd_r, shd_g, shd_b, shd_a,
           btlr_red, btlr_green, btlr_blue, btlr_alpha,
           type, width);
  return ppr_.w_tag();
}

// [[Rcpp::export]]
std::string a_ppr(std::string text_align,
                  int pb, int pt, int pl, int pr,
                  int shd_r, int shd_g, int shd_b, int shd_a,
                  IntegerVector btlr_red, IntegerVector btlr_green,
                  IntegerVector btlr_blue, IntegerVector btlr_alpha,
                  CharacterVector type, IntegerVector width ) {
  ppr ppr_(text_align, pb, pt, pl, pr,
           shd_r, shd_g, shd_b, shd_a,
           btlr_red, btlr_green, btlr_blue, btlr_alpha,
           type, width);
  return ppr_.a_tag();
}

// [[Rcpp::export]]
std::string css_ppr(std::string text_align,
                  int pb, int pt, int pl, int pr,
                  int shd_r, int shd_g, int shd_b, int shd_a,
                  IntegerVector btlr_red, IntegerVector btlr_green,
                  IntegerVector btlr_blue, IntegerVector btlr_alpha,
                  CharacterVector type, IntegerVector width ) {
  ppr ppr_(text_align, pb, pt, pl, pr,
           shd_r, shd_g, shd_b, shd_a,
           btlr_red, btlr_green, btlr_blue, btlr_alpha,
           type, width);
  return ppr_.css();
}
